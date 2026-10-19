"""월간 성과 보고서 계산: 월별 KPI · 매출 구성 · 주차 추이 · 급변 옵션(채널 이동/광고 연결)."""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

from core.analysis import delta, pct, won

from .data import AdRaw, InsightData, OptionSheet


@dataclass
class MonthlySettings:
    account: str
    written: str = ""
    cutoff: date | None = None        # None → 방문자가 있는 마지막 날 (당일 부분집계 제외)
    min_diff: float = 500_000         # 급변: |증감액| ≥ min_diff AND |증감률| ≥ min_pct
    min_pct: float = 15
    shift_min: float = 100_000        # 채널 이동: 상대 채널이 이 금액 이상 &
    shift_ratio: float = 0.2          #           내 변동액의 20% 이상 반대로 움직임
    recent_days: int = 7              # 광고 상태: 광고 raw 마지막 N일 광고비 > 0 → 집행중
    manual_ads: list = field(default_factory=list)     # [(date, 광고비, 광고전환매출)] 대시보드 누락 보정
    channel_pairs: list = field(default_factory=list)  # [(옵션ID, 옵션ID)] 옵션명이 달라 자동 매칭 안 되는 쌍
    headline: str | None = None       # 포커스 배너 문구 (None이면 자동)
    headline_sub: str | None = None


def _month_end(y: int, m: int) -> date:
    return date(y, m, calendar.monthrange(y, m)[1])


def _days(s: date, e: date):
    return [s + timedelta(i) for i in range((e - s).days + 1)]


def man(v: float) -> str:
    """만원 단위 축약 (배너·차트 라벨용)."""
    return f"{v / 10000:,.0f}만"


# ---------------- 기간 집계 ----------------

def resolve_periods(raw: AdRaw | None, daily_dates) -> None:
    """기간 합계 광고 파일의 시작·종료일 확정 (이미 정해진 건 유지). 연도는 판매 데이터에서 해당 월이 있는 해."""
    if not raw:
        return
    years = {}
    for d in daily_dates:
        years.setdefault(d.month, d.year)
    for p in raw.periods:
        if p.s and p.e or not p.md:
            continue
        m1, d1, m2, d2 = p.md
        y1 = years.get(m1, max(years.values()) if years else date.today().year)
        y2 = y1 + (1 if m2 < m1 else 0)
        try:
            p.s = date(y1, m1, d1)
            p.e = _month_end(y2, m2) if d2 == 0 else date(y2, m2, d2)
        except ValueError:
            p.s = p.e = None


def active_periods(raw: AdRaw | None, total_only: bool = False) -> list:
    if not raw:
        return []
    return [p for p in raw.periods if p.s and p.e and (p.use_total or not total_only)]


def build_ad_rows(dash: list | None, manual: list, raw: AdRaw | None, cutoff: date) -> list[dict]:
    """광고 합계 행. 우선순위: 광고보고서 기간 합계(계정 전체) > 대시보드 > 수기 보정 > 광고 raw 일합계.
    기간 합계 파일은 쿠팡이 직접 낸 기간 전체 수치라 수기 입력 대시보드보다 우선하고, 겹치는 대시보드 행은 버린다."""
    rows, covered = [], set()
    for p in active_periods(raw, total_only=True):
        days = set(_days(p.s, min(p.e, cutoff)))
        if p.e <= cutoff and not days & covered:
            rows.append({"s": p.s, "e": p.e, "cost": p.cost, "adrev": p.rev, "src": f"광고보고서({p.source})"})
            covered |= days
    by_period = set(covered)
    for r in dash or []:  # 대시보드는 같은 주에 캠페인별 여러 행 → 기간 합계 파일과 겹칠 때만 버린다
        days = set(_days(r["s"], r["e"]))
        if not days & by_period:
            rows.append({**r, "src": "대시보드"})
            covered |= days
    for d, c, a in manual:
        if d not in covered:
            rows.append({"s": d, "e": d, "cost": c, "adrev": a, "src": "수기"}); covered.add(d)
    if raw:
        for d, items in sorted(raw.by_day.items()):
            if d not in covered and d <= cutoff:
                rows.append({"s": d, "e": d, "cost": sum(x[1] for x in items),
                             "adrev": sum(x[2] for x in items), "src": "광고raw"})
    return rows


def cvr_of(daily: dict, days: list) -> float:
    """구매전환율(%) = 주문 ÷ 조회 — Wing·서플라이어 허브 화면 정의와 같고 월 총계와 정확히 일치.
    조회 칸이 없는 파일이면 일별 구매전환율의 방문자 가중평균으로 대신한다."""
    view = sum(daily[d].get("view", 0) for d in days)
    if view:
        return sum(daily[d]["ord"] for d in days) / view * 100
    vis = sum(daily[d]["vis"] for d in days)
    return sum(daily[d]["vis"] * daily[d]["cvr"] for d in days) / vis if vis else 0.0


def period(daily: dict, ad_rows: list, s: date, e: date) -> dict:
    days = [d for d in _days(s, e) if d in daily]
    t = {"s": s, "e": e, "days": len(days),
         "rev": sum(daily[d]["rev"] for d in days), "vis": sum(daily[d]["vis"] for d in days),
         "ord": sum(daily[d]["ord"] for d in days)}
    t["view"] = sum(daily[d].get("view", 0) for d in days)
    t["cvr"] = cvr_of(daily, days)
    rows = [a for a in ad_rows if a["s"] >= s and a["e"] <= e]
    covered = {d for a in rows for d in _days(a["s"], a["e"])}
    t["addays"] = len(covered)
    t["ad_ok"] = bool(rows) and days and sum(d in covered for d in days) >= 0.9 * len(days)
    t["cost"] = sum(a["cost"] for a in rows)
    t["adrev"] = sum(a["adrev"] for a in rows)
    t["ad_missing"] = [d for d in days if d not in covered]
    t["ad_filled"] = [(a["s"], a["e"], a["src"]) for a in rows if a["src"] != "대시보드"]
    t["natural"] = t["rev"] - t["adrev"] if t["ad_ok"] else None
    t["ad_ratio"] = t["cost"] / t["rev"] * 100 if t["ad_ok"] and t["rev"] else None
    t["ad_share"] = t["adrev"] / t["rev"] * 100 if t["ad_ok"] and t["rev"] else None
    return t


# ---------------- 옵션 · 채널 · 광고 ----------------

def pick_option_sheet(cands: list[OptionSheet], daily_total: float) -> tuple[OptionSheet | None, float | None]:
    """같은 월 상품별 시트가 여러 개면 일별 매출 합계와 가장 가까운 것."""
    if not cands:
        return None, None
    best = min(cands, key=lambda o: abs(o.total - daily_total))
    gap = (best.total / daily_total - 1) * 100 if daily_total else None
    return best, gap


def name_key(name: str) -> str:
    s = re.sub(r"\[[^\]]*\]", "", name.lower())   # [전국설치] 같은 머리말
    s = re.sub(r"\b[a-z]{1,3}_", "", s)            # JK_ / VZ_ 같은 옵션 접두어
    return re.sub(r"[^0-9a-z가-힣]", "", s)


def channel_groups(opt_dicts: list[dict], pairs: list) -> dict:
    """옵션ID → 그룹키. 정규화한 옵션명이 같고 판매방식이 다른 옵션끼리 묶는다."""
    by_key: dict = {}
    for od in opt_dicts:
        for oid, o in od.items():
            by_key.setdefault(name_key(o["name"]), {})[oid] = o["ch"]
    group = {}
    for k, members in by_key.items():
        if len(set(members.values())) > 1:
            for oid in members:
                group[oid] = k
    for a, b in pairs:
        a, b = str(a).strip(), str(b).strip()
        k = group.get(a) or group.get(b) or f"manual_{a}"
        group[a] = group[b] = k
    return group


def ad_context(raw: AdRaw | None, cutoff: date, recent_days: int) -> dict:
    """옵션별 광고비·광고 상태. 일별 raw + 기간 합계 파일(그 기간 안에서는 기간 합계가 우선)."""
    opt_cost: dict = {}
    if raw:
        for d, items in raw.by_day.items():
            if d > cutoff:
                continue
            for oid, c, _ in items:
                opt_cost[(oid, d)] = opt_cost.get((oid, d), 0) + c
    periods = [p for p in active_periods(raw) if p.s <= cutoff]
    dated_days = sorted({d for _, d in opt_cost})
    period_days = {d for p in periods for d in _days(p.s, min(p.e, cutoff))}
    day_set = set(dated_days) | period_days
    all_days = sorted(day_set)
    recent = set(all_days[-recent_days:])
    by_opt: dict = {}
    for (o, d), c in opt_cost.items():
        if c > 0:
            by_opt.setdefault(o, set()).add(d)

    def spend(oid, s, e):
        inside = [p for p in periods if p.s >= s and p.e <= e]
        covered = {d for p in inside for d in _days(p.s, p.e)}
        total = sum(p.opt.get(oid, [0, 0])[0] for p in inside)
        return total + sum(c for (o, d), c in opt_cost.items() if o == oid and s <= d <= e and d not in covered)

    def coverage(s, e):
        inside = {d for p in periods if p.s >= s and p.e <= e for d in _days(p.s, p.e)}
        return len(inside | {d for d in dated_days if s <= d <= e}), (e - s).days + 1

    def gaps(s, e):
        out = []
        for d in _days(s, e):
            if d in day_set:
                continue
            if out and d - out[-1][1] == timedelta(1):
                out[-1][1] = d
            else:
                out.append([d, d])
        return out

    def status(oid):
        if not all_days:
            return "-"
        ds = by_opt.get(oid, set())
        if ds & recent:
            return "집행중"
        spent_periods = [p for p in periods if p.opt.get(oid, [0])[0] > 0]
        for p in spent_periods:
            if set(_days(p.s, min(p.e, cutoff))) & recent:
                return f"집행(기간합계 {p.s.month}/{p.s.day}~{p.e.month}/{p.e.day})"
        last_d = max(ds) if ds else None
        last_p = max(spent_periods, key=lambda p: p.e) if spent_periods else None
        if last_d is None and last_p is None:
            return "미집행"
        if last_p and (last_d is None or last_p.e >= last_d):
            last, label = last_p.e, f"{last_p.s.month}/{last_p.s.day}~{last_p.e.month}/{last_p.e.day} 집행 후"
        else:
            last, label = last_d, f"마지막 {last_d.month}/{last_d.day}"
        if last + timedelta(1) not in day_set:
            return f"확인필요({last.month}/{last.day} 이후 광고 데이터 없음)"
        return f"중단({label})"

    return {"has": bool(all_days), "spend": spend, "coverage": coverage, "gaps": gaps, "status": status,
            "recent": (min(recent), max(recent)) if recent else None}


def movers(a: dict, b: dict, scale_b: float, window: tuple, group: dict, ad: dict,
           pa: tuple, pb: tuple, s: MonthlySettings) -> list[dict]:
    w0, w1, w_sc = window
    rev = lambda od, oid, sc=1.0: od.get(oid, {}).get("rev", 0) * sc
    members: dict = {}
    for oid, k in group.items():
        members.setdefault(k, []).append(oid)
    ch_of = lambda oid: (b.get(oid) or a.get(oid) or w0.get(oid) or w1.get(oid) or {}).get("ch")
    out = []
    for oid in set(a) | set(b):
        ra, rb = rev(a, oid), rev(b, oid, scale_b)
        diff = rb - ra
        pct_ = diff / ra * 100 if ra else None
        if abs(diff) < s.min_diff or (pct_ is not None and abs(pct_) < s.min_pct):
            continue
        info = b.get(oid) or a.get(oid)
        note, shift, net = "", "", None
        others = [o for o in members.get(group.get(oid), []) if ch_of(o) != info["ch"]]
        if others:
            enough = lambda od: od * diff < 0 and abs(od) >= max(s.shift_min, s.shift_ratio * abs(diff))
            same = sum(rev(b, o, scale_b) - rev(a, o) for o in others)
            lag = sum(rev(w1, o, w_sc) - rev(w0, o) for o in others)
            kind = "동시" if enough(same) else ("시차" if enough(lag) else "")
            if kind:
                other_ch = ch_of(others[0])
                shift = f"{info['ch']}→{other_ch}" if diff < 0 else f"{other_ch}→{info['ch']}"
                od = same if kind == "동시" else lag
                net = diff + od
                verdict = ("채널 이동으로 설명됨(실질 급변 아님)" if abs(net) < s.min_diff
                           else f"이동분 빼도 순{'감소' if net < 0 else '증가'} {won(abs(net))} · 별도 원인 확인")
                note = f"{shift} {kind} 이동 추정 · {other_ch} 옵션 {od:+,.0f}원 · {verdict}"
        out.append({"oid": oid, "ch": info["ch"], "name": info["name"], "a": ra, "b": rb, "diff": diff,
                    "pct": pct_, "va": a.get(oid, {}).get("vis", 0), "vb": b.get(oid, {}).get("vis", 0) * scale_b,
                    "shift": shift, "note": note, "net": net,
                    "ad_status": ad["status"](oid) if ad["has"] else "-",
                    "ad_a": ad["spend"](oid, *pa) if ad["has"] else None,
                    "ad_b": ad["spend"](oid, *pb) if ad["has"] else None})
    return sorted(out, key=lambda x: x["diff"])


# ---------------- 주차 추이 ----------------

def weekly(daily: dict, start: date, cutoff: date, max_weeks: int = 8) -> list[dict]:
    """월요일 시작 주 단위 일평균. 4일 미만 주는 partial."""
    s = start - timedelta(start.weekday())
    weeks = []
    while s <= cutoff:
        e = min(s + timedelta(6), cutoff)
        days = [d for d in _days(max(s, start), e) if d in daily]
        if days:
            vis = sum(daily[d]["vis"] for d in days)
            weeks.append({"s": days[0], "e": days[-1], "n": len(days), "partial": len(days) < 4,
                          "rev": sum(daily[d]["rev"] for d in days) / len(days),
                          "vis": vis / len(days),
                          "cvr": cvr_of(daily, days),
                          "label": f"{days[0].month}/{days[0].day}~{days[-1].month}/{days[-1].day}"})
        s += timedelta(7)
    return weeks[-max_weeks:]


def trend_word(vals: list[float]) -> str:
    if len(vals) < 3:
        return "판단 보류"
    a, b, c = vals[-3:]
    if a < b < c:
        return "상승 중"
    if a > b > c:
        return "하락 중"
    return "상승 전환" if c > b else ("하락 전환" if c < b else "보합")


# ---------------- 컨텍스트 ----------------

def build_context(ins: InsightData, raw: AdRaw | None, dash: list | None, s: MonthlySettings) -> dict:
    daily = ins.daily
    if not daily:
        raise ValueError("일별 판매 데이터('기간별 판매량' 시트)가 없습니다.")
    cutoff = s.cutoff or max(d for d, v in daily.items() if v["vis"] > 0)
    daily = {d: v for d, v in daily.items() if d <= cutoff}
    resolve_periods(raw, daily)
    ad_rows = build_ad_rows(dash, s.manual_ads, raw, cutoff)

    months = []
    for y, m in sorted({(d.year, d.month) for d in daily}):
        start, end = date(y, m, 1), min(_month_end(y, m), cutoff)
        partial = end < _month_end(y, m)
        t = period(daily, ad_rows, start, end)
        t.update(y=y, m=m, partial=partial, month_days=_month_end(y, m).day,
                 label=f"{m}월" + (f"(1~{end.day}일)" if partial else ""), short=f"{m}월",
                 tag=f"{m}/1~{end.day}" if partial else f"{m}월")
        months.append(t)
    last, prev = months[-1], (months[-2] if len(months) >= 2 else None)

    # 진행 중인 달은 직전 달 같은 일수와 비교
    same = None
    if prev and last["partial"]:
        pe = min(date(prev["y"], prev["m"], 1) + timedelta(last["e"].day - 1), prev["e"])
        same = period(daily, ad_rows, date(prev["y"], prev["m"], 1), pe)
        same["label"] = f"{prev['m']}월 동기간(1~{pe.day}일)"

    def cmp(key, kind="pct", per_day=False):
        """(증감, 설명). 진행 중인 달: 매출·방문자·전환율은 동기간, 광고는 일평균 vs 직전 달."""
        if not prev:
            return None, ""
        if key in ("cost",):
            if not (last["ad_ok"] and prev["ad_ok"]):
                return None, "광고 데이터 부족"
            if last["partial"]:
                return delta(last["cost"] / last["days"], prev["cost"] / prev["days"]), f"일평균 기준, {prev['short']} 대비"
            return delta(last["cost"], prev["cost"]), f"{prev['short']} 대비"
        if key == "ad_ratio":
            if last["ad_ratio"] is None or prev["ad_ratio"] is None:
                return None, "광고 데이터 부족"
            return delta(last["ad_ratio"], prev["ad_ratio"], "pp"), f"{prev['short']} 대비 · 낮을수록 효율적"
        base = same if same else prev
        return delta(last[key], base[key], kind), (f"{same['label']} 대비" if same else f"{prev['short']} 대비")

    def rows_of(key, fmt):
        return [{"tag": t["tag"], "text": fmt(t[key]) if t.get(key) is not None else None,
                 "cls": "cur" if t is last else "prev"} for t in months]

    kpis = []
    for label, key, fmt, kind in [("총매출", "rev", won, "pct"), ("광고비", "cost", won, "pct"),
                                  ("매출 대비 광고비 비율", "ad_ratio", lambda v: pct(v, 1), "pp"),
                                  ("방문자 수", "vis", lambda v: f"{v:,.0f}명", "pct"),
                                  ("구매전환율", "cvr", lambda v: pct(v, 2), "pp")]:
        rows = rows_of(key, fmt)
        if key == "cost":
            for r, t in zip(rows, months):
                r["text"] = fmt(t["cost"]) if t["ad_ok"] else None
        d, how = cmp(key, kind)
        kpis.append({"label": label, "rows": rows, "delta": d, "how": how})

    # 옵션 · 급변
    checks, option_by_month = [], {}
    for t in months:
        cands = [o for o in ins.options if o.month == t["m"]]
        best, gap = pick_option_sheet(cands, t["rev"])
        if best is None:
            checks.append(("warn", f"{t['short']} 상품별 판매량이 없어 이 달이 포함된 급변 옵션 분석을 건너뜁니다."))
            continue
        option_by_month[t["m"]] = (best, gap)
        tol = 10 if t["partial"] else 2  # 진행 중인 달은 당일 부분집계가 섞일 수 있음
        if best.partial:
            checks.append(("info", f"{t['short']} 상품별 데이터는 상위 상품만 담긴 시트({best.source})라 급변 상품도 그 범위 안에서만 비교합니다."))
        elif gap is not None and abs(gap) > tol:
            checks.append(("warn", f"{t['short']} 상품별 매출 합계({won(best.total)})가 일별 합계({won(t['rev'])})와 "
                                   f"{gap:+.0f}% 차이 납니다. 조회 기간이 맞는지 확인해 주세요. ({best.source})"))
    ad = ad_context(raw, cutoff, s.recent_days)
    opt_dicts = [option_by_month[t["m"]][0].rows for t in months if t["m"] in option_by_month]
    group = channel_groups(opt_dicts, s.channel_pairs)

    def scale_of(t):
        if not t["partial"] or t["m"] not in option_by_month:
            return 1.0
        sheet = option_by_month[t["m"]][0]
        if sheet.partial:  # 상위 상품만 있는 시트는 합계 대조가 불가 → 단순 일수 환산
            return t["month_days"] / t["days"] if t["days"] else 1.0
        # 전체 시트: 당일 부분집계가 섞일 수 있어 일별 기준 월환산 매출에 비례 배분
        return (t["rev"] / t["days"] * t["month_days"]) / sheet.total if sheet.total and t["days"] else 1.0

    win_months = [t for t in months if t["m"] in option_by_month]
    window = ((option_by_month[win_months[0]["m"]][0].rows, option_by_month[win_months[-1]["m"]][0].rows,
               scale_of(win_months[-1])) if win_months else ({}, {}, 1.0))
    mover_sets = []
    for ta, tb in zip(months, months[1:]):
        if ta["m"] not in option_by_month or tb["m"] not in option_by_month:
            continue
        pa, pb = (ta["s"], ta["e"]), (tb["s"], tb["e"])
        mv = movers(option_by_month[ta["m"]][0].rows, option_by_month[tb["m"]][0].rows, scale_of(tb),
                    window, group, ad, pa, pb, s)
        ca, cb = ad["coverage"](*pa) if ad["has"] else (0, 0), ad["coverage"](*pb) if ad["has"] else (0, 0)
        mover_sets.append({"id": f"mv{ta['m']}_{tb['m']}", "title": f"{ta['short']} → {tb['short']}",
                           "a_label": ta["short"], "b_label": tb["short"] + (" 월환산" if tb["partial"] else ""),
                           "rows": mv, "n_down": sum(1 for m in mv if m["diff"] < 0),
                           "n_up": sum(1 for m in mv if m["diff"] > 0),
                           "n_shift": sum(1 for m in mv if m["shift"]),
                           "cov_a": ca, "cov_b": cb, "partial": tb["partial"],
                           "top_only": option_by_month[ta["m"]][0].partial or option_by_month[tb["m"]][0].partial,
                           "gaps": ad["gaps"](ta["s"], tb["e"]) if ad["has"] else []})

    # 데이터 점검
    for t in months:
        if not t["ad_ok"]:
            checks.append(("warn", f"{t['short']} 광고 합계 데이터가 부족해(커버 {t['addays']}/{t['days']}일) 광고 지표를 비워둡니다."))
        elif t["ad_missing"]:
            miss = ", ".join(f"{d.month}/{d.day}" for d in t["ad_missing"])
            checks.append(("info", f"{t['short']} 광고 데이터 누락일: {miss} (광고 지표가 소폭 과소)"))
        for d, e2, src in t["ad_filled"]:
            rng = f"{d.month}/{d.day}" + (f"~{e2.month}/{e2.day}" if e2 != d else "")
            checks.append(("info", f"{rng} 광고 합계는 {src} 값을 사용했습니다."))
    if not ad["has"]:
        checks.append(("info", "광고 raw가 없어 급변 옵션의 광고 상태·광고비 열은 비워둡니다."))
    for p in (raw.periods if raw else []):
        if not (p.s and p.e):
            checks.append(("warn", f"'{p.source}'은(는) 날짜 없는 기간 합계 파일인데 기간을 알 수 없어 제외했습니다. 화면에서 기간을 지정하세요."))
            continue
        tag = f"'{p.source}' ({p.s.month}/{p.s.day}~{p.e.month}/{p.e.day} 기간 합계, 캠페인 {len(p.campaigns)}개, 광고비 {won(p.cost)})"
        if not p.use_total:
            checks.append(("info", f"{tag}는 계정 전체가 아니라고 보고 옵션별 광고비에만 반영했습니다."))
        if dash:
            inside = [r for r in dash if r["s"] >= p.s and r["e"] <= p.e]
            if inside and p.use_total:
                dsum = sum(r["cost"] for r in inside)
                gap = (dsum / p.cost - 1) * 100 if p.cost else 0
                if abs(gap) > 1:
                    checks.append(("warn", f"{tag}와 대시보드 같은 기간 광고비({won(dsum)})가 {gap:+.1f}% 다릅니다. 광고보고서 값을 사용합니다."))
    checks.extend(("warn", n) for n in ins.notes)

    # 주차 추이 · 배너
    weeks = weekly(daily, months[0]["s"] if len(months) < 3 else months[-2]["s"], cutoff)
    full = [w for w in weeks if not w["partial"]]
    word = trend_word([w["rev"] for w in full])
    headline, sub = auto_headline(months, same, full, word)

    return {"s": s, "cutoff": cutoff, "months": months, "last": last, "prev": prev, "same": same,
            "kpis": kpis, "weeks": weeks, "trend_word": word,
            "headline": s.headline or headline, "headline_sub": s.headline_sub or sub,
            "auto_headline": headline, "auto_headline_sub": sub,
            "movers": mover_sets, "ad": ad, "checks": checks,
            "period_label": f"{months[0]['y']}년 {months[0]['m']}월~{last['m']}월"
                            + (f" ({last['m']}/{last['e'].day}까지)" if last["partial"] else "")}


def auto_headline(months, same, weeks, word) -> tuple[str, str]:
    last = months[-1]
    prev = months[-2] if len(months) >= 2 else None
    if not prev:
        return f"{last['label']} 매출 {won(last['rev'])}", ""
    base = same or prev
    d = delta(last["rev"], base["rev"])
    parts = [f"{last['label']} 매출은 {base.get('label', prev['short'])} 대비 {d['text'] if d else '-'}"]
    if len(weeks) >= 3:
        w = [man(x["rev"]) for x in weeks[-3:]]
        parts.append(f"주차별 일평균은 {' → '.join(w)}원으로 {word}")
    head = ", ".join(parts)
    sub = []
    if last["ad_ratio"] is not None and prev["ad_ratio"] is not None:
        sub.append(f"매출 대비 광고비 {prev['ad_ratio']:.1f}% → {last['ad_ratio']:.1f}%")
    if last["ad_share"] is not None and prev["ad_share"] is not None:
        sub.append(f"자연매출 비중 {100 - prev['ad_share']:.0f}% → {100 - last['ad_share']:.0f}%")
    sub.append(f"구매전환율 {prev['cvr']:.2f}% → {last['cvr']:.2f}%")
    return head, " · ".join(sub)


def report_data(ctx: dict) -> dict:
    """통합 탭용 요약 수치 (보고서 HTML에 내장)."""
    keep = ("rev", "vis", "view", "ord", "cvr", "cost", "adrev", "ad_ok", "days", "partial", "label", "short", "tag", "m")
    def slim(t):
        return {k: t[k] for k in keep if k in t} | {"cvr_w": t["cvr"] * t["vis"]}
    return {"kind": "monthly", "version": 1, "account": ctx["s"].account,
            "months": [slim(t) for t in ctx["months"]],
            "same": slim(ctx["same"]) if ctx["same"] else None,
            "trend_word": ctx["trend_word"], "period_label": ctx["period_label"],
            "n_movers": sum(len(m["rows"]) for m in ctx["movers"])}
