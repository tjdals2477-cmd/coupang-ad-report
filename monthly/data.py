"""월간 성과 보고서 입력 파일 읽기.

- 쿠팡 Wing 비즈니스 인사이트: '기간별 판매량'(일별) + '상품별 판매량'(옵션별) 시트. 상품별만 따로 받은 파일도 가능.
- 쿠팡 광고 raw: 일자 × 광고집행 옵션ID 행 (광고센터 보고서 또는 광고보고서 xlsm의 'raw' 시트)
- 광고 주차 대시보드(선택): 시작일·종료일·광고비용·광고매출 행 (시트 = 계정)
시트·컬럼은 이름으로 찾으므로 순서가 바뀌어도 된다.
"""
from __future__ import annotations

import io
import re
import warnings
from dataclasses import dataclass, field
from datetime import date, datetime

import openpyxl

warnings.filterwarnings("ignore", module="openpyxl")


class MonthlyFileError(ValueError):
    pass


def num(v) -> float:
    if v in (None, "", "-"):
        return 0.0
    try:
        return float(str(v).replace(",", "").replace("%", "").strip())
    except ValueError:
        return 0.0


def to_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.match(r"\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})", v)
        if m:
            return date(int(m[1]), int(m[2]), int(m[3]))
    return None


def _month_from_name(name: str) -> int | None:
    m = re.search(r"(\d{1,2})\s*월", name)
    return int(m[1]) if m and 1 <= int(m[1]) <= 12 else None


def _open(data: bytes):
    try:
        return openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as e:  # 손상·비엑셀 파일
        raise MonthlyFileError(f"엑셀 파일을 열 수 없습니다: {e}") from e


def _header(row) -> list[str]:
    return [str(c).strip() if c is not None else "" for c in row]


# ---------------- 비즈니스 인사이트 ----------------

DAILY_COLS = {"날짜", "방문자", "주문", "구매전환율", "매출(원)"}
OPTION_COLS = {"옵션 ID", "옵션명", "판매방식", "매출(원)", "방문자"}


@dataclass
class OptionSheet:
    source: str            # "파일명 · 시트명"
    month: int | None
    rows: dict             # {옵션ID: {name, ch, rev, vis, ord}}

    @property
    def total(self) -> float:
        return sum(r["rev"] for r in self.rows.values())


@dataclass
class InsightData:
    daily: dict = field(default_factory=dict)     # {date: {vis, ord, cvr, rev}}
    options: list = field(default_factory=list)   # [OptionSheet]
    notes: list = field(default_factory=list)


def read_insight(data: bytes, name: str, into: InsightData | None = None) -> InsightData:
    out = into or InsightData()
    wb = _open(data)
    found = False
    daily_month = None
    pending_opts = []
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        try:
            hdr = _header(next(it))
        except StopIteration:
            continue
        if DAILY_COLS <= set(hdr):
            ix = {h: hdr.index(h) for h in DAILY_COLS}
            months = []
            for r in it:
                d = to_date(r[ix["날짜"]])
                if not d:
                    continue  # 하단 총계 행
                out.daily[d] = {"vis": num(r[ix["방문자"]]), "ord": num(r[ix["주문"]]),
                                "cvr": num(r[ix["구매전환율"]]), "rev": num(r[ix["매출(원)"]])}
                months.append(d.month)
            if months:
                daily_month = max(set(months), key=months.count)
                found = True
        elif OPTION_COLS <= set(hdr):
            ix = {h: hdr.index(h) for h in OPTION_COLS}
            i_ord = hdr.index("주문") if "주문" in hdr else None
            rows = {}
            for r in it:
                oid = r[ix["옵션 ID"]]
                if oid in (None, ""):
                    continue
                oid = str(oid).split(".")[0].strip()
                rows[oid] = {"name": str(r[ix["옵션명"]] or ""), "ch": str(r[ix["판매방식"]] or ""),
                             "rev": num(r[ix["매출(원)"]]), "vis": num(r[ix["방문자"]]),
                             "ord": num(r[i_ord]) if i_ord is not None else 0.0}
            pending_opts.append((ws.title, rows))
            found = True
    for title, rows in pending_opts:
        month = daily_month or _month_from_name(name)
        if month is None:
            out.notes.append(f"'{name}' 상품별 시트의 월을 알 수 없어 제외했습니다. 파일명에 'n월'을 넣어주세요.")
            continue
        out.options.append(OptionSheet(source=f"{name} · {title}", month=month, rows=rows))
    if not found:
        raise MonthlyFileError(f"'{name}'에서 '기간별 판매량'(날짜·방문자·매출) 또는 '상품별 판매량'(옵션 ID·판매방식·매출) 시트를 찾지 못했습니다.")
    return out


# ---------------- 광고 raw (옵션 단위 일별) ----------------

AD_RAW_COLS = ("날짜", "광고집행 옵션ID", "광고비")
AD_REV_COLS = ("총 전환매출액(14일)", "총 전환매출액 (14일)(원)")


@dataclass
class AdRaw:
    by_day: dict = field(default_factory=dict)   # {date: [(옵션ID, 광고비, 14일 전환매출)]}


def read_ad_raw(data: bytes, name: str, into: AdRaw | None = None) -> AdRaw:
    """같은 날짜가 여러 파일에 있으면 나중에 읽은 파일이 덮어쓴다 (업로드 순서 = 최신 순으로 올리면 됨)."""
    out = into or AdRaw()
    wb = _open(data)
    sheets = [wb["raw"]] if "raw" in wb.sheetnames else wb.worksheets
    for ws in sheets:
        it = ws.iter_rows(values_only=True)
        try:
            hdr = _header(next(it))
        except StopIteration:
            continue
        if not all(c in hdr for c in AD_RAW_COLS):
            continue
        i_d, i_o, i_c = (hdr.index(c) for c in AD_RAW_COLS)
        i_r = next((hdr.index(c) for c in AD_REV_COLS if c in hdr), None)
        rows: dict = {}
        for r in it:
            d = to_date(r[i_d])
            if not d:
                continue
            rows.setdefault(d, []).append((str(r[i_o]).split(".")[0], num(r[i_c]), num(r[i_r]) if i_r is not None else 0.0))
        out.by_day.update(rows)
        return out
    raise MonthlyFileError(f"'{name}'에서 광고 raw(날짜·광고집행 옵션ID·광고비 컬럼)를 찾지 못했습니다.")


# ---------------- 광고 주차 대시보드 ----------------

def dashboard_sheets(data: bytes) -> list[str]:
    return _open(data).sheetnames


def read_ad_dashboard(data: bytes, sheet: str) -> list[dict]:
    """[{s, e, cost, adrev}] — '시작일','종료일','광고비용','광고매출' 헤더가 있는 표."""
    wb = _open(data)
    if sheet not in wb.sheetnames:
        raise MonthlyFileError(f"대시보드에 '{sheet}' 시트가 없습니다.")
    rows = list(wb[sheet].iter_rows(values_only=True))
    for i, r in enumerate(rows[:15]):
        hdr = _header(r)
        if {"시작일", "종료일", "광고비용", "광고매출"} <= set(hdr):
            ix = {h: hdr.index(h) for h in ("시작일", "종료일", "광고비용", "광고매출")}
            out = []
            for r2 in rows[i + 1:]:
                s, e = to_date(r2[ix["시작일"]]), to_date(r2[ix["종료일"]])
                if s and e:
                    out.append({"s": s, "e": e, "cost": num(r2[ix["광고비용"]]), "adrev": num(r2[ix["광고매출"]])})
            return out
    raise MonthlyFileError(f"'{sheet}' 시트에서 시작일·종료일·광고비용·광고매출 헤더를 찾지 못했습니다.")
