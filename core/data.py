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
        return _read_report_sheet(file)


# 광고보고서 xlsm 'raw' 시트·광고센터 상품 전환 보고서 등은 같은 값을 조금 다른 이름으로 쓴다 (공백·'(원)' 무시)
_COL_ALIASES = {
    "캠페인ID": "캠페인 ID", "캠페인명": "캠페인명", "캠페인": "캠페인명", "캠페인이름": "캠페인명",
    "광고집행상품명": "광고집행 상품명", "광고집행옵션ID": "광고집행 옵션ID",
    "노출수": "노출수", "클릭수": "클릭수", "광고비": "광고비",
    "총주문수(14일)": "총 주문수(14일)", "총주문수(1일)": "총 주문수(1일)",
    "총전환매출액(14일)": "총 전환매출액(14일)", "총전환매출액(1일)": "총 전환매출액(1일)",
}


def _canon(col) -> str:
    c = str(col).strip()
    return _COL_ALIASES.get(re.sub(r"\s+", "", c).replace("(원)", ""), c)


def _read_report_sheet(file) -> pd.DataFrame:
    """여러 시트 파일(광고보고서 xlsm 등)이면 'raw' 시트 → 옵션ID·광고비가 있는 첫 시트 순으로 고른다."""
    xl = pd.ExcelFile(file)
    names = xl.sheet_names
    order = (["raw"] if "raw" in names else []) + [n for n in names if n != "raw"]
    for sheet in order:
        cols = {_canon(c) for c in xl.parse(sheet, nrows=0).columns}
        if {"광고집행 옵션ID", "광고비"} <= cols:
            return xl.parse(sheet)
    return xl.parse(order[0])


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

    df.columns = [_canon(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    if "캠페인 ID" not in df.columns and "캠페인명" in df.columns:
        df["캠페인 ID"] = df["캠페인명"]  # 상품 전환 보고서에는 캠페인 ID가 없음 → 캠페인명으로 매칭
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
    df["date"] = _detect_dates(df)
    return df


DATE_COLS = ["날짜", "일자", "date", "Date"]


def _parse_dates(s: pd.Series) -> pd.Series:
    """20260901 / '2026-09-01' / '2026.09.01' / 엑셀 날짜 모두 처리."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return s.dt.normalize()
    text = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    compact = text.str.fullmatch(r"\d{8}")
    out = pd.to_datetime(text.where(~compact), errors="coerce", format="mixed")
    out = out.fillna(pd.to_datetime(text.where(compact), errors="coerce", format="%Y%m%d"))
    return out.dt.normalize()


def _detect_dates(df: pd.DataFrame) -> pd.Series:
    """'일별'로 내려받은 보고서의 날짜 컬럼. 없거나 대부분 해석되지 않으면 NaT."""
    for col in DATE_COLS:
        if col in df.columns:
            parsed = _parse_dates(df[col])
            if len(parsed) and parsed.notna().mean() >= 0.9:
                return parsed
    return pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")


def date_span(df: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp] | None:
    """보고서에 날짜가 있으면 (시작일, 종료일)."""
    d = df["date"].dropna() if "date" in df.columns else pd.Series(dtype="datetime64[ns]")
    return (d.min(), d.max()) if len(d) else None


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


_RANGE_RE = re.compile(
    r"(\d{1,2})\s*[월./]\s*(\d{1,2})\s*일?\s*[~\-–]\s*(?:(\d{1,2})\s*[월./]\s*)?(\d{1,2})\s*일?")


def guess_meta(filename: str) -> dict:
    """'안트_9월 광고보고서.xlsx' → {'account': '안트', 'month': 9, 'range': None}
    '9월16~9월22일.xlsx' → {'month': 9, 'range': ((9, 16), (9, 22)), 'label': '9/16~22', 'days': 7}"""
    stem = re.sub(r"\.(xlsx|xlsm|xls|csv)$", "", filename, flags=re.I)
    account = re.split(r"[_\s]", stem, maxsplit=1)[0] if "_" in stem else ""
    r = _RANGE_RE.search(stem)
    if r:
        m1, d1 = int(r.group(1)), int(r.group(2))
        m2, d2 = int(r.group(3) or m1), int(r.group(4))
        if 1 <= m1 <= 12 and 1 <= m2 <= 12 and 1 <= d1 <= 31 and 1 <= d2 <= 31:
            year = pd.Timestamp.today().year
            start = pd.Timestamp(year, m1, d1)
            end = pd.Timestamp(year + (1 if m2 < m1 else 0), m2, d2)
            label = f"{m1}/{d1}~{d2}" if m1 == m2 else f"{m1}/{d1}~{m2}/{d2}"
            return {"account": account, "month": m1, "range": ((m1, d1), (m2, d2)),
                    "label": label, "days": (end - start).days + 1}
    m = re.search(r"(\d{1,2})\s*월", stem)
    month = int(m.group(1)) if m and 1 <= int(m.group(1)) <= 12 else None
    return {"account": account, "month": month, "range": None, "label": None, "days": None}
