"""쿠팡 Wing 광고 보고서(엑셀/CSV) 로딩과 집계."""
from __future__ import annotations

import io
import re
import warnings

import pandas as pd

# 전환 기준(1일/14일)에 따라 사용할 컬럼
WINDOW_COLS = {
    "14일": {"orders": "총 주문수(14일)", "revenue": "총 전환매출액(14일)"},
    "1일": {"orders": "총 주문수(1일)", "revenue": "총 전환매출액(1일)"},
}

BASE_COLS = ["캠페인 ID", "캠페인명", "광고집행 상품명", "광고집행 옵션ID", "노출수", "클릭수", "광고비"]
NUMERIC_COLS = ["노출수", "클릭수", "광고비"] + [c for w in WINDOW_COLS.values() for c in w.values()]


class ReportFileError(ValueError):
    """업로드 파일이 쿠팡 광고 보고서 형식이 아닐 때."""


def _read_raw(file, filename: str) -> pd.DataFrame:
    name = filename.lower()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # openpyxl "no default style" 경고
        if name.endswith(".csv"):
            raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
            for enc in ("utf-8-sig", "cp949"):
                try:
                    return pd.read_csv(io.BytesIO(raw), encoding=enc)
                except UnicodeDecodeError:
                    continue
            raise ReportFileError("CSV 인코딩을 읽을 수 없습니다 (UTF-8 / CP949 아님).")
        return pd.read_excel(file)


def _to_number(s: pd.Series) -> pd.Series:
    if not pd.api.types.is_numeric_dtype(s):  # pandas 3의 'str' dtype과 object 모두 처리
        s = s.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce").fillna(0)


def load_report(file, filename: str, window: str = "14일") -> pd.DataFrame:
    """보고서를 읽어 표준 컬럼(cost/revenue/orders/clicks/impressions)을 붙여 반환."""
    try:
        df = _read_raw(file, filename)
    except ReportFileError:
        raise
    except Exception as e:  # 손상 파일, 비엑셀 파일 등
        raise ReportFileError(f"'{filename}' 파일을 열 수 없습니다: {e}") from e

    df.columns = [str(c).strip() for c in df.columns]
    need = BASE_COLS + list(WINDOW_COLS[window].values())
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ReportFileError(
            f"'{filename}'에 필수 컬럼이 없습니다: {', '.join(missing)}\n"
            "쿠팡 Wing > 광고 보고서에서 내려받은 원본 파일인지 확인해 주세요."
        )

    for c in NUMERIC_COLS:
        if c in df.columns:
            df[c] = _to_number(df[c])

    df["캠페인 ID"] = df["캠페인 ID"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["캠페인명"] = df["캠페인명"].fillna("-").astype(str)
    df["광고집행 상품명"] = df["광고집행 상품명"].fillna("-").astype(str)
    df["광고집행 옵션ID"] = df["광고집행 옵션ID"].astype(str).str.replace(r"\.0$", "", regex=True)

    w = WINDOW_COLS[window]
    df["cost"] = df["광고비"]
    df["revenue"] = df[w["revenue"]]
    df["orders"] = df[w["orders"]]
    df["clicks"] = df["클릭수"]
    df["impressions"] = df["노출수"]
    return df


def _ratio(num, den, scale=100.0):
    """0으로 나누기를 0으로 처리하는 비율 계산 (스칼라·Series 모두 지원)."""
    if isinstance(den, pd.Series):
        return (num / den.where(den != 0)).fillna(0) * scale
    return (num / den * scale) if den else 0.0


def summarize(df: pd.DataFrame) -> dict:
    cost, rev = float(df["cost"].sum()), float(df["revenue"].sum())
    clicks, orders = float(df["clicks"].sum()), float(df["orders"].sum())
    return {
        "cost": cost,
        "revenue": rev,
        "clicks": clicks,
        "orders": orders,
        "impressions": float(df["impressions"].sum()),
        "roas": _ratio(rev, cost),
        "cvr": _ratio(orders, clicks),
    }


def _with_rates(g: pd.DataFrame) -> pd.DataFrame:
    g["roas"] = _ratio(g["revenue"], g["cost"])
    g["cvr"] = _ratio(g["orders"], g["clicks"])
    return g


_AGG = dict(cost=("cost", "sum"), revenue=("revenue", "sum"), orders=("orders", "sum"),
            clicks=("clicks", "sum"), impressions=("impressions", "sum"))


def campaign_table(df: pd.DataFrame) -> pd.DataFrame:
    """캠페인 ID 단위 집계. 캠페인명은 파일 내 가장 많이 쓰인 이름."""
    g = df.groupby("캠페인 ID").agg(**_AGG)
    names = df.groupby("캠페인 ID")["캠페인명"].agg(lambda s: s.mode().iat[0])
    g.insert(0, "name", names)
    g = _with_rates(g).reset_index().rename(columns={"캠페인 ID": "id"})
    return g.sort_values("cost", ascending=False).reset_index(drop=True)


def clean_product_name(name: str, limit: int = 60) -> str:
    """'상품명,{"색상":"..."},5개,5개,주광색,주광색' → '상품명 · 5개 · 주광색' (JSON 옵션 제거, 중복 옵션 제거)."""
    parts, seen = [], set()
    for part in re.sub(r"\{[^{}]*\}", "", str(name)).split(","):
        part = part.strip()
        if part and part not in seen:
            seen.add(part)
            parts.append(part)
    text = " · ".join(parts) or str(name)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def product_table(df: pd.DataFrame, campaign_id: str) -> pd.DataFrame:
    """특정 캠페인의 광고집행 상품(옵션ID) 단위 집계. 배너 등 상품 없는 행 제외."""
    sub = df[(df["캠페인 ID"] == campaign_id) & (df["광고집행 상품명"] != "-")]
    if sub.empty:
        return pd.DataFrame(columns=["option_id", "name", "full_name", *_AGG.keys(), "roas", "cvr"])
    g = sub.groupby("광고집행 옵션ID").agg(full_name=("광고집행 상품명", "first"), **_AGG)
    g = _with_rates(g).reset_index().rename(columns={"광고집행 옵션ID": "option_id"})
    g.insert(1, "name", g["full_name"].map(clean_product_name))
    return g


def guess_meta(filename: str) -> dict:
    """'안트_9월 광고보고서.xlsx' → {'account': '안트', 'month': 9}."""
    stem = re.sub(r"\.(xlsx|xls|csv)$", "", filename, flags=re.I)
    m = re.search(r"(\d{1,2})\s*월", stem)
    month = int(m.group(1)) if m and 1 <= int(m.group(1)) <= 12 else None
    account = re.split(r"[_\s]", stem, maxsplit=1)[0] if "_" in stem else ""
    return {"account": account, "month": month}
