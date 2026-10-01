"""보고서 섹션 ①~⑤에 들어갈 값을 계산한다. 렌더링과 분리해 테스트 가능하게 유지."""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

import pandas as pd

from .data import campaign_table, date_span, product_table, summarize


@dataclass
class Settings:
    account: str
    prev_label: str = "직전월"
    cur_label: str = "금월"
    prev_days: int = 0          # 0이면 일평균 미표시
    cur_days: int = 0
    prev_total_sales: float = 0  # 0이면 "입력 필요"
    cur_total_sales: float = 0
    window: str = "14일"
    min_cost: float = 50_000     # ③ 최고 효율 후보 최소 광고비
    min_clicks: int = 10         # ⑤ CVR 순위 최소 클릭수
    min_product_cost: float = 10_000  # ⑤ ROAS 순위 최소 상품 광고비 (소액 착시 방지)
    top_n: int = 5
    written: str = ""
    # {캠페인 ID: {"target_roas": float|None, "daily_budget": float|None}}
    targets: dict = field(default_factory=dict)
    # 일별 데이터일 때 구간별 총매출 {구간 key: 금액} (key는 make_buckets의 key)
    bucket_sales: dict = field(default_factory=dict)


# ---------- 공통 포맷/증감 ----------

def won(v: float) -> str:
    return f"{round(v):,}원"


def pct(v: float, d: int = 2) -> str:
    return f"{v:,.{d}f}%"


def delta(cur: float, prev: float, kind: str = "pct") -> dict | None:
    """kind='pct' → 상대 증감률(%), 'pp' → 차이(%p). 증가=up(빨강▲), 감소=down(파랑▼)."""
    if kind == "pct":
        if not prev:
            return None
        diff = (cur - prev) / prev * 100
        unit = "%"
    else:
        diff = cur - prev
        unit = "%p"
    if abs(diff) < 0.005:
        return {"dir": "flat", "text": f"– 0.0{unit}", "value": 0.0}
    arrow, d = ("▲", "up") if diff > 0 else ("▼", "down")
    return {"dir": d, "text": f"{arrow} {abs(diff):,.1f}{unit}", "value": diff}


def _metric_rows(prev: dict | None, cur: dict) -> list[dict]:
    """②③ 비교 카드의 4개 지표 행."""
    spec = [("집행 광고비", "cost", won, "pct"), ("광고 전환 매출", "revenue", won, "pct"),
            ("광고수익률(ROAS)", "roas", pct, "pp"), ("전환율(CVR)", "cvr", pct, "pp")]
    rows = []
    for label, key, fmt, kind in spec:
        rows.append({
            "label": label,
            "prev": fmt(prev[key]) if prev else "-",
            "cur": fmt(cur[key]),
            "delta": delta(cur[key], prev[key], kind) if prev else None,
        })
    return rows


def _camp_dict(row: pd.Series) -> dict:
    return {k: (float(row[k]) if k not in ("id", "name") else str(row[k]))
            for k in ("id", "name", "cost", "revenue", "orders", "clicks", "roas", "cvr")}


# ---------- 섹션 ① ----------

def kpi_cards(p: dict, c: dict, s: Settings) -> list[dict]:
    def daily(v, days):
        return won(v / days) if days else None

    cards = []
    for label, pv, cv, fmt, kind in [
        ("총 광고집행비용", p["cost"], c["cost"], won, "pct"),
        ("총매출", s.prev_total_sales, s.cur_total_sales, won, "pct"),
        ("광고 전환 매출", p["revenue"], c["revenue"], won, "pct"),
        ("ROAS", p["roas"], c["roas"], pct, "pp"),
        ("전환율(CVR)", p["cvr"], c["cvr"], pct, "pp"),
    ]:
        is_sales = label == "총매출"
        card = {
            "label": label,
            "prev": fmt(pv) if (pv or not is_sales) else None,   # None → "입력 필요"
            "cur": fmt(cv) if (cv or not is_sales) else None,
            "prev_raw": pv, "cur_raw": cv,
            "decimals": 2 if kind == "pp" else 0,
            "suffix": "%" if kind == "pp" else "원",
            "delta": delta(cv, pv, kind) if (pv and cv) or not is_sales else None,
            "note": "",
        }
        if kind == "pct" and s.prev_days and s.cur_days and pv and cv:
            card["note"] = f"일평균 {s.prev_label} {daily(pv, s.prev_days)}<br>{s.cur_label} {daily(cv, s.cur_days)}"
        cards.append(card)

    # 보조 설명: ROAS/CVR 카드
    cards[3]["note"] = f"광고전환매출 ÷ 광고비 ({s.window} 전환 기준)"
    cards[4]["note"] = f"주문수 ÷ 클릭수 ({round(c['clicks']):,}클릭 · {round(c['orders']):,}건)"
    if s.cur_total_sales and c["revenue"]:
        cards[1]["note"] = (cards[1]["note"] + "<br>" if cards[1]["note"] else "") + \
            f"{s.cur_label} 광고매출 비중 {c['revenue'] / s.cur_total_sales * 100:.1f}%"
    return cards


# ---------- 섹션 ① 추이 차트 (총매출 · 광고전환매출 · 집행 광고비) ----------

def short_won(v: float) -> str:
    """차트 라벨용 축약: 90,989,740 → '9,099만', 120,000,000 → '1.2억'."""
    if v >= 1e8:
        return f"{v / 1e8:,.2f}".rstrip("0").rstrip(".") + "억"
    if v >= 1e4:
        return f"{v / 1e4:,.0f}만"
    return f"{v:,.0f}"


def _nice_step(raw: float) -> float:
    if raw <= 0:
        return 1.0
    mag = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * mag:
            return m * mag
    return 10 * mag


def _spread(ys: list[float], gap: float, lo: float, hi: float) -> list[float]:
    """라벨 y좌표가 gap보다 가까우면 위아래로 벌린다 (입력 순서 유지)."""
    order = sorted(range(len(ys)), key=lambda i: ys[i])
    out = list(ys)
    for a, b in zip(order, order[1:]):
        if out[b] - out[a] < gap:
            out[b] = out[a] + gap
    overflow = out[order[-1]] - hi
    if overflow > 0:
        for i in order:
            out[i] -= overflow
    return [max(lo, y) for y in out]


DAILY_MAX_DAYS = 14  # 이 기간 이하이면 1일 단위, 초과하면 7일 단위


def choose_granularity(prev_df: pd.DataFrame, cur_df: pd.DataFrame) -> str | None:
    """두 보고서 모두 날짜가 있을 때만 'day' 또는 'week'. 둘 다 같은 단위를 써야 크기 비교가 맞다."""
    spans = [date_span(prev_df), date_span(cur_df)]
    if not all(spans):
        return None
    longest = max((e - st).days + 1 for st, e in spans)
    return "day" if longest <= DAILY_MAX_DAYS else "week"


def make_buckets(df: pd.DataFrame, side: str, granularity: str) -> list[dict]:
    """보고서를 1일/7일 구간으로 나눠 광고비·광고전환매출 합계. 데이터가 없는 날도 0으로 채운다."""
    start, end = date_span(df)
    daily = df.dropna(subset=["date"]).groupby("date")[["cost", "revenue"]].sum()
    daily = daily.reindex(pd.date_range(start, end, freq="D"), fill_value=0)
    size = 1 if granularity == "day" else 7
    buckets = []
    for i in range(0, len(daily), size):
        chunk = daily.iloc[i:i + size]
        b0, b1 = chunk.index[0], chunk.index[-1]
        label = f"{b0.month}/{b0.day}" if b0 == b1 else (
            f"{b0.month}/{b0.day}~{b1.day}" if b0.month == b1.month else f"{b0.month}/{b0.day}~{b1.month}/{b1.day}")
        buckets.append({"key": f"{side}:{b0:%Y-%m-%d}", "start": b0, "end": b1, "label": label,
                        "days": len(chunk), "cost": float(chunk["cost"].sum()),
                        "revenue": float(chunk["revenue"].sum())})
    return buckets


SERIES_SPEC = [("sales", "총매출", "#5B4FE5", "circle"),
               ("revenue", "광고 전환 매출", "#14b8a6", "square"),
               ("cost", "집행 광고비", "#FF7A45", "diamond")]


def trend_points(prev_df: pd.DataFrame, cur_df: pd.DataFrame, p: dict, c: dict, s: Settings) -> dict | None:
    """차트에 찍을 점 목록. 총매출·광고전환매출·광고비 중 하나라도 비면 None."""
    gran = choose_granularity(prev_df, cur_df)
    groups = []
    if gran:
        for side, df, label in (("prev", prev_df, s.prev_label), ("cur", cur_df, s.cur_label)):
            pts = make_buckets(df, side, gran)
            for b in pts:
                b["sales"] = s.bucket_sales.get(b["key"])
            groups.append({"label": label, "points": pts})
    else:  # 월 합계 파일: 직전월·금월 2점
        groups = [{"label": s.prev_label, "points": [{"key": "prev", "label": s.prev_label, "sales": s.prev_total_sales,
                                                       "revenue": p["revenue"], "cost": p["cost"]}]},
                  {"label": s.cur_label, "points": [{"key": "cur", "label": s.cur_label, "sales": s.cur_total_sales,
                                                      "revenue": c["revenue"], "cost": c["cost"]}]}]
    for g in groups:
        for pt in g["points"]:
            if any(pt.get(k) in (None, "") or float(pt[k]) <= 0 for k in ("sales", "revenue", "cost")):
                return None
    return {"granularity": gran or "month", "groups": groups}


def trend_chart(prev_df: pd.DataFrame, cur_df: pd.DataFrame, p: dict, c: dict, s: Settings) -> dict | None:
    data = trend_points(prev_df, cur_df, p, c, s)
    if not data:
        return None
    groups = data["groups"]
    flat = [(gi, pt) for gi, g in enumerate(groups) for pt in g["points"]]
    n = len(flat)

    W, H, L, R, T, B = 760, 300, 64, 150, 34, 44
    pw, ph = W - L - R, H - T - B
    # x: 구간 사이 간격 1칸, 직전월/금월 사이에는 1칸 더 띄워 구분
    slots = [i + gi for i, (gi, _) in enumerate(flat)]
    span = max(slots[-1], 1)
    pad = pw * (0.18 if n == 2 else 0.04)
    xs = [L + pad + (pw - 2 * pad) * (sl / span) for sl in slots] if n > 1 else [L + pw / 2]
    top = max(float(pt[k]) for _, pt in flat for k in ("sales", "revenue", "cost"))
    step = _nice_step(top * 1.1 / 4)
    ymax = step * math.ceil(top * 1.1 / step)

    def y(v):
        return T + ph * (1 - v / ymax)

    series = []
    for key, label, color, shape in SERIES_SPEC:
        vals = [float(pt[key]) for _, pt in flat]
        pts = [(x, y(v)) for x, v in zip(xs, vals)]
        # 직전월과 금월 사이는 선을 끊는다 (기간이 이어지지 않음)
        segs = []
        for gi in range(len(groups)):
            idx = [i for i, (g, _) in enumerate(flat) if g == gi]
            if len(idx) > 1:
                segs.append(" ".join(f"{pts[i][0]:.1f},{pts[i][1]:.1f}" for i in idx))
        series.append({"key": key, "label": label, "color": color, "shape": shape, "values": vals,
                       "pts": pts, "segs": segs, "end_txt": short_won(vals[-1]),
                       "connect": n == 2})
    # 끝점 직접 라벨 (겹침 방지)
    ly = _spread([sr["pts"][-1][1] + 4 for sr in series], 16, T + 8, T + ph)
    for sr, yy in zip(series, ly):
        sr["end_label_y"] = yy
    if n == 2:  # 월 2점: 시작점 값도 표시
        ly0 = _spread([sr["pts"][0][1] + 4 for sr in series], 16, T + 8, T + ph)
        for sr, yy in zip(series, ly0):
            sr["start_label_y"] = yy
            sr["start_txt"] = short_won(sr["values"][0])

    every = 1 if n <= 16 else math.ceil(n / 16)  # x축 라벨 솎기 (기간별로 처음·끝은 항상 표시)
    shown = set()
    for gi in range(len(groups)):
        idx = [i for i, (g, _) in enumerate(flat) if g == gi]
        pick = idx[::every]
        if pick[-1] != idx[-1]:
            if idx[-1] - pick[-1] < every:  # 끝 라벨과 너무 가까우면 앞 라벨을 뺀다
                pick = pick[:-1]
            pick.append(idx[-1])
        shown.update(pick)
    cols = []
    for i, (gi, pt) in enumerate(flat):
        half = (xs[1] - xs[0]) / 2 if n > 1 else 40
        partial = data["granularity"] == "week" and pt.get("days", 7) < 7  # 7일 미만 구간 (월말 등)
        cols.append({"x": xs[i], "label": pt["label"] + (f" ({pt['days']}일)" if partial else ""), "partial": partial,
                     "label_main": pt["label"], "label_sub": f"({pt['days']}일)" if partial else "",
                     "show": i in shown,
                     "hit_x": xs[i] - min(half, 60), "hit_w": min(half, 60) * 2, "group": groups[gi]["label"]})
    dividers, headers = [], []
    if data["granularity"] != "month":
        for gi, g in enumerate(groups):
            idx = [i for i, (g2, _) in enumerate(flat) if g2 == gi]
            headers.append({"x": (xs[idx[0]] + xs[idx[-1]]) / 2, "label": g["label"]})
            if gi > 0:
                dividers.append((xs[idx[0] - 1] + xs[idx[0]]) / 2)
    ticks = [{"y": y(v), "label": short_won(v) if v else "0"}
             for v in [step * i for i in range(int(round(ymax / step)) + 1)]]
    unit = {"day": "1일 단위", "week": "7일 단위", "month": "월 합계"}[data["granularity"]]
    tooltip = []
    for i, (gi, pt) in enumerate(flat):
        title = f"{groups[gi]['label']} · {cols[i]['label']}" if data["granularity"] != "month" else pt["label"]
        tooltip.append({"title": title, "values": [won(sr["values"][i]) for sr in series],
                        "deltas": [delta(sr["values"][i], sr["values"][i - 1]) if i > 0 and flat[i - 1][0] == gi
                                   or (n == 2 and i == 1) else None for sr in series]})
    return {"w": W, "h": H, "left": L, "right": W - R, "top": T, "bottom": T + ph, "n": n,
            "series": series, "cols": cols, "ticks": ticks, "dividers": dividers, "headers": headers,
            "granularity": data["granularity"], "unit": unit, "tooltip": tooltip,
            "has_partial": any(c["partial"] for c in cols)}


# ---------- 섹션 ② 최대 이슈 ----------

def issue_campaign(prev_c: pd.DataFrame, cur_c: pd.DataFrame, cur_total_cost: float) -> dict | None:
    """금월 광고비 1위 캠페인 — 단, 직전월에도 집행(광고비>0)된 캠페인 중에서."""
    prev_ids = set(prev_c.loc[prev_c["cost"] > 0, "id"])
    cand = cur_c[(cur_c["cost"] > 0) & cur_c["id"].isin(prev_ids)]
    if cand.empty:
        return None
    cur_row = cand.iloc[0]  # cur_c는 광고비 내림차순
    prev_row = prev_c[prev_c["id"] == cur_row["id"]].iloc[0]
    cur, prev = _camp_dict(cur_row), _camp_dict(prev_row)
    return {
        "cur": cur, "prev": prev,
        "rows": _metric_rows(prev, cur),
        "share": cur["cost"] / cur_total_cost * 100 if cur_total_cost else 0,
        "renamed": prev["name"] != cur["name"],
    }


# ---------- 섹션 ③ 최고 효율 ----------

def _best_by_roas(c: pd.DataFrame, min_cost: float) -> pd.Series | None:
    pool = c[(c["cost"] >= max(min_cost, 1)) & (c["revenue"] > 0)]
    if pool.empty:  # 기준 금액 이상 캠페인이 없으면 광고비>0 전체에서
        pool = c[(c["cost"] > 0) & (c["revenue"] > 0)]
    if pool.empty:
        return None
    return pool.sort_values(["roas", "revenue"], ascending=False).iloc[0]


def best_campaign(prev_c: pd.DataFrame, cur_c: pd.DataFrame, min_cost: float) -> dict | None:
    cur_row = _best_by_roas(cur_c, min_cost)
    if cur_row is None:
        return None
    same = prev_c[(prev_c["id"] == cur_row["id"]) & (prev_c["cost"] > 0)]
    if not same.empty:
        mode, prev_row = "same", same.iloc[0]
    else:
        mode, prev_row = "fallback", _best_by_roas(prev_c, min_cost)
    cur = _camp_dict(cur_row)
    prev = _camp_dict(prev_row) if prev_row is not None else None
    return {"cur": cur, "prev": prev, "mode": mode, "rows": _metric_rows(prev, cur)}


# ---------- 섹션 ⑤ 캠페인 현황 + TOP 상품 ----------

def _top_products(df: pd.DataFrame, cid: str, s: Settings) -> dict:
    p = product_table(df, cid)
    spent = p[p["cost"] > 0]
    by_roas = spent[spent["cost"] >= s.min_product_cost].sort_values(["roas", "revenue"], ascending=False).head(s.top_n)
    by_cvr = spent[spent["clicks"] >= s.min_clicks].sort_values(["cvr", "orders"], ascending=False).head(s.top_n)

    def rows(t):
        return [{"name": r["name"], "full_name": r["full_name"], "cost": won(r["cost"]),
                 "revenue": won(r["revenue"]), "clicks": f"{int(r['clicks']):,}",
                 "orders": f"{int(r['orders']):,}", "roas": pct(r["roas"]), "cvr": pct(r["cvr"])}
                for _, r in t.iterrows()]

    return {"roas": rows(by_roas), "cvr": rows(by_cvr), "n_products": len(spent),
            "n_converting": int((spent["orders"] > 0).sum())}


def campaign_cards(df: pd.DataFrame, prev_c: pd.DataFrame, cur_c: pd.DataFrame, s: Settings) -> list[dict]:
    prev_ids = set(prev_c.loc[prev_c["cost"] > 0, "id"])
    cards = []
    for i, r in cur_c[cur_c["cost"] > 0].reset_index(drop=True).iterrows():
        t = s.targets.get(str(r["id"]), {}) or {}
        target, budget = t.get("target_roas"), t.get("daily_budget")
        target = float(target) if target not in (None, "") and float(target) > 0 else None
        budget = float(budget) if budget not in (None, "") and float(budget) > 0 else None
        achieve = r["roas"] / target * 100 if target else None
        cards.append({
            "idx": i, "id": str(r["id"]), "name": str(r["name"]),
            "is_new": str(r["id"]) not in prev_ids,
            "target": pct(target, 0) if target else None,
            "budget": won(budget) if budget else None,
            "cost": won(r["cost"]), "revenue": won(r["revenue"]),
            "roas": pct(r["roas"]), "cvr": pct(r["cvr"]),
            "achieve": achieve, "hit": achieve is not None and achieve >= 100,
            "fill": min(achieve, 100) if achieve is not None else 0,
            "top": _top_products(df, str(r["id"]), s),
        })
    return cards


# ---------- 섹션 ④ 3단계 요약 ----------

def short(name: str, limit: int = 20) -> str:
    """요약용 캠페인명 축약: '[메모]', '(09.16)' 같은 접두 괄호 제거, '_' → 공백."""
    text = re.sub(r"^\s*(\[[^\]]*\]|\([^)]*\))\s*", "", name).replace("_", " ").strip() or name
    return text if len(text) <= limit else text[: limit - 1] + "…"


def summary_steps(p: dict, c: dict, issue: dict | None, best: dict | None,
                  cur_c: pd.DataFrame, cards: list[dict], s: Settings) -> list[dict]:
    """규칙 기반 초안 — 한 줄 한 사실, 줄당 짧게. 앱에서 사용자가 수정할 수 있다.
    줄 규칙: '• ' 불릿, '→ ' 결론. ▲/▼ 증감은 보고서에서 빨강/파랑으로 표시된다."""
    def d_text(d):
        return d["text"] if d else "-"

    # STEP 1 — 성과 진단
    dc, dr = delta(c["cost"], p["cost"]), delta(c["revenue"], p["revenue"])
    droas = delta(c["roas"], p["roas"], "pp")
    up_rev, up_roas = bool(dr and dr["value"] > 0), c["roas"] >= p["roas"]
    verdict = {(True, True): "규모·효율 동반 개선", (True, False): "규모 확대, 효율 하락",
               (False, True): "효율 개선, 규모 축소", (False, False): "규모·효율 동반 하락"}[(up_rev, up_roas)]
    step1 = "\n".join([
        f"• 광고비 {d_text(dc)}",
        f"• 전환매출 {d_text(dr)}",
        f"• ROAS {pct(c['roas'], 0)} ({d_text(droas)})",
        f"→ {verdict}",
    ])

    # STEP 2 — 핵심 이슈 (최대 3줄)
    parts = []
    if issue:
        parts.append(f"• {short(issue['cur']['name'])} 비중 {issue['share']:.0f}%, "
                     f"ROAS {d_text(issue['rows'][2]['delta'])}")
    miss = [k for k in cards if k["achieve"] is not None and not k["hit"]]
    if miss:
        parts.append(f"• 목표 ROAS 미달 {len(miss)}개")
    else:
        low = cur_c[(cur_c["cost"] >= s.min_cost) & (cur_c["roas"] < c["roas"])]
        if not low.empty:
            parts.append(f"• 평균 ROAS 미만 {len(low)}개 캠페인")
    leak = cur_c[(cur_c["cost"] > 0) & (cur_c["revenue"] <= 0)]
    if not leak.empty:
        parts.append(f"• 전환 0건 {len(leak)}개 ({won(leak['cost'].sum())})")
    step2 = "\n".join(parts) or "• 특이 이슈 없음"

    # STEP 3 — 개선 방향 (최대 3줄)
    acts = []
    if best:
        acts.append(f"• 확대: {short(best['cur']['name'])} ({pct(best['cur']['roas'], 0)})")
    weak = cur_c[cur_c["cost"] >= s.min_cost].sort_values("roas").head(2)
    weak = weak[weak["roas"] < c["roas"]]
    if not weak.empty:
        more = f" 외 {len(weak) - 1}개" if len(weak) > 1 else ""
        acts.append(f"• 축소: {short(weak.iloc[0]['name'])} ({pct(weak.iloc[0]['roas'], 0)}){more}")
    acts.append("• 정리: 전환 0건 상품 제외")
    step3 = "\n".join(acts)

    return [
        {"title": "성과 진단", "text": step1},
        {"title": "핵심 이슈", "text": step2},
        {"title": "개선 방향", "text": step3},
    ]


def headline(issue: dict | None, c: dict, cur_c: pd.DataFrame) -> dict:
    top2 = cur_c.head(2)
    share = top2["cost"].sum() / c["cost"] * 100 if c["cost"] else 0
    if issue:
        title = f"금월 광고비의 {issue['share']:.1f}%가 '{issue['cur']['name']}' 1개 캠페인에 집중"
    else:
        title = f"상위 2개 캠페인이 금월 광고비의 {share:.1f}% 차지"
    return {"title": title, "text": f"집행 캠페인 {int((cur_c['cost'] > 0).sum())}개 · 상위 2개 캠페인 광고비 비중 {share:.1f}%"}


# ---------- 전체 ----------

def build_context(prev_df: pd.DataFrame, cur_df: pd.DataFrame, s: Settings) -> dict:
    p, c = summarize(prev_df), summarize(cur_df)
    prev_c, cur_c = campaign_table(prev_df), campaign_table(cur_df)
    issue = issue_campaign(prev_c, cur_c, c["cost"])
    best = best_campaign(prev_c, cur_c, s.min_cost)
    cards = campaign_cards(cur_df, prev_c, cur_c, s)
    return {
        "s": s,
        "prev": p, "cur": c,
        "kpis": kpi_cards(p, c, s),
        "trend": trend_chart(prev_df, cur_df, p, c, s),
        "headline": headline(issue, c, cur_c),
        "issue": issue,
        "best": best,
        "steps": summary_steps(p, c, issue, best, cur_c, cards, s),
        "campaigns": cards,
        "n_campaigns": int((cur_c["cost"] > 0).sum()),
        "campaign_list": cur_c[["id", "name", "cost"]].to_dict("records"),
    }
