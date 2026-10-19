"""월간 성과 보고서 HTML 렌더링 · 계정 통합 탭 · 탭형 병합."""
from __future__ import annotations

from core.analysis import delta, pct, won
from core.render import _env, extract_report_data, safe_json

from .analysis import man, report_data


def composition_chart(months: list[dict]) -> dict | None:
    """월별 100% 누적 막대: 자연매출(아래) + 광고전환매출(위)."""
    cols = [t for t in months if t["ad_ok"]]
    if not cols:
        return None
    W, H, top, bottom, left = 640, 260, 30, 40, 20
    ph = H - top - bottom
    bw = min(110, (W - left * 2) / len(cols) * 0.5)
    step = (W - left * 2) / len(cols)
    bars = []
    for i, t in enumerate(cols):
        x = left + step * i + (step - bw) / 2
        nat = max(0.0, 100 - t["ad_share"])
        h_nat = ph * nat / 100
        bars.append({"x": x, "w": bw, "cx": x + bw / 2, "label": t["label"],
                     "nat_y": top + ph - h_nat, "nat_h": h_nat, "ad_y": top, "ad_h": ph - h_nat,
                     "nat": nat, "ad": t["ad_share"], "rev": man(t["rev"]) + "원"})
    return {"w": W, "h": H, "bottom": top + ph, "bars": bars}


def weekly_chart(weeks: list[dict]) -> dict | None:
    """주차별 일평균 매출(막대) + 구매전환율(선)."""
    if len(weeks) < 2:
        return None
    W, H, top, bottom, left, right = 760, 280, 36, 46, 16, 16
    ph = H - top - bottom
    step = (W - left - right) / len(weeks)
    bw = step * 0.48
    vmax = (max(w["rev"] for w in weeks) or 1) / 0.6   # 막대는 아래 60%, 선은 위쪽 띠 → 라벨 겹침 방지
    cmin, cmax = min(w["cvr"] for w in weeks), max(w["cvr"] for w in weeks)
    span = (cmax - cmin) or 1
    pts = []
    for i, w in enumerate(weeks):
        cx = left + step * i + step / 2
        h = ph * w["rev"] / vmax
        cy = top + 10 + (ph * 0.2) * (1 - (w["cvr"] - cmin) / span)
        pts.append({"cx": cx, "x": cx - bw / 2, "w": bw, "y": top + ph - h, "h": h, "cy": cy,
                    "rev": man(w["rev"]), "cvr": f"{w['cvr']:.2f}%", "label": w["label"],
                    "partial": w["partial"], "n": w["n"]})
    path = " ".join(("M" if i == 0 else "L") + f"{p['cx']:.1f} {p['cy']:.1f}" for i, p in enumerate(pts))
    return {"w": W, "h": H, "bottom": top + ph, "pts": pts, "path": path}


def mover_view(ms: dict) -> dict:
    def ad_txt(v):
        return "-" if v is None else won(v)
    rows = []
    for m in ms["rows"]:
        rows.append({**m, "a_t": won(m["a"]), "b_t": won(m["b"]),
                     "diff_t": f"{m['diff']:+,.0f}원", "pct_t": (f"{m['pct']:+.0f}%" if m["pct"] is not None else "신규"),
                     "dir": "up" if m["diff"] > 0 else "down",
                     "ad_a_t": ad_txt(m["ad_a"]), "ad_b_t": ad_txt(m["ad_b"]),
                     "ad_d": delta(m["ad_b"], m["ad_a"]) if m["ad_a"] else None,
                     "st_cls": ("st-hit" if m["ad_status"] == "집행중" else
                                "st-miss" if m["ad_status"].startswith("중단") else
                                "st-new" if m["ad_status"].startswith("확인") else "st-none")})
    gap_txt = ", ".join(f"{a.month}/{a.day}" + (f"~{b.month}/{b.day}" if b != a else "") for a, b in ms["gaps"])
    return {**ms, "rows": rows, "gap_txt": gap_txt}


def render_monthly(ctx: dict) -> str:
    tpl = _env.get_template("monthly_report.html.j2")
    return tpl.render(**ctx, comp=composition_chart(ctx["months"]), wchart=weekly_chart(ctx["weeks"]),
                      mover_views=[mover_view(m) for m in ctx["movers"]],
                      data_json=safe_json(report_data(ctx)))


# ---------------- 계정 통합 ----------------

def render_monthly_summary(datas: list[dict], title: str) -> str:
    """각 계정 보고서에 내장된 월별 수치를 월 단위로 합산."""
    labels = [t["short"] for t in datas[0]["months"]]
    months = []
    for i, short in enumerate(labels):
        ts = [next((t for t in d["months"] if t["short"] == short), None) for d in datas]
        ts = [t for t in ts if t]
        ad_ok = all(t["ad_ok"] for t in ts) and len(ts) == len(datas)
        vis = sum(t["vis"] for t in ts)
        rev = sum(t["rev"] for t in ts)
        cost, adrev = sum(t["cost"] for t in ts), sum(t["adrev"] for t in ts)
        months.append({"short": short, "label": ts[0]["label"], "tag": ts[0].get("tag", ts[0]["label"]), "partial": ts[0]["partial"],
                       "rev": rev, "vis": vis, "cvr": sum(t["cvr_w"] for t in ts) / vis if vis else 0,
                       "cost": cost if ad_ok else None, "adrev": adrev if ad_ok else None,
                       "ad_ratio": cost / rev * 100 if ad_ok and rev else None,
                       "ad_share": adrev / rev * 100 if ad_ok and rev else None, "n": len(ts)})
    last, prev = months[-1], (months[-2] if len(months) >= 2 else None)
    sames = [d.get("same") for d in datas]
    same_rev = sum(x["rev"] for x in sames) if all(sames) else None
    same_vis = sum(x["vis"] for x in sames) if all(sames) else None
    base_how = "직전 달 동기간 대비" if last["partial"] and same_rev else "직전 달 대비"
    rows = []
    for d in datas:
        t = d["months"][-1]
        base = d.get("same") or (d["months"][-2] if len(d["months"]) >= 2 else None)
        rows.append({"account": d["account"], "rev": won(t["rev"]),
                     "rev_d": delta(t["rev"], base["rev"]) if base else None,
                     "cost": won(t["cost"]) if t["ad_ok"] else "-",
                     "ratio": pct(t["cost"] / t["rev"] * 100, 1) if t["ad_ok"] and t["rev"] else "-",
                     "share": pct(100 - t["adrev"] / t["rev"] * 100, 0) if t["ad_ok"] and t["rev"] else "-",
                     "vis": f"{t['vis']:,.0f}", "cvr": pct(t["cvr"], 2), "trend": d.get("trend_word", "-"),
                     "rev_share": t["rev"] / last["rev"] * 100 if last["rev"] else 0})
    kpis = [
        {"label": "총매출", "rows": [(m["tag"], won(m["rev"])) for m in months],
         "delta": delta(last["rev"], same_rev if (last["partial"] and same_rev) else (prev or {}).get("rev", 0)),
         "how": base_how},
        {"label": "광고비", "rows": [(m["tag"], won(m["cost"]) if m["cost"] is not None else "-") for m in months],
         "delta": None, "how": ""},
        {"label": "매출 대비 광고비 비율", "rows": [(m["tag"], pct(m["ad_ratio"], 1) if m["ad_ratio"] is not None else "-") for m in months],
         "delta": delta(last["ad_ratio"], prev["ad_ratio"], "pp") if prev and last["ad_ratio"] is not None and prev["ad_ratio"] is not None else None,
         "how": "낮을수록 효율적"},
        {"label": "방문자 수", "rows": [(m["tag"], f"{m['vis']:,.0f}명") for m in months],
         "delta": delta(last["vis"], same_vis if (last["partial"] and same_vis) else (prev or {}).get("vis", 0)), "how": base_how},
        {"label": "구매전환율", "rows": [(m["tag"], pct(m["cvr"], 2)) for m in months],
         "delta": delta(last["cvr"], prev["cvr"], "pp") if prev else None, "how": "방문자 가중평균"},
    ]
    tpl = _env.get_template("monthly_summary.html.j2")
    return tpl.render(title=title, kpis=kpis, rows=rows, months=months,
                      comp=composition_chart([{**m, "ad_ok": m["ad_share"] is not None, "ad_share": m["ad_share"] or 0} for m in months]),
                      accounts=[d["account"] for d in datas], period_label=datas[0].get("period_label", ""))


def merge_monthly(reports: list[dict], title: str, add_summary: bool = True) -> tuple[str, list[str]]:
    items = [{"id": f"r{i}", "label": r["label"], "html": r["html"]} for i, r in enumerate(reports)]
    skipped, datas = [], []
    for r in reports:
        d = extract_report_data(r["html"])
        (datas.append(d) if d and d.get("kind") == "monthly" else skipped.append(r["label"]))
    if add_summary and len(datas) >= 2:
        items.append({"id": "summary", "label": "계정 통합 보고서",
                      "html": render_monthly_summary(datas, f"{title} · 계정 통합")})
    embedded = {it["id"]: {"label": it["label"], "html": it["html"]} for it in items}
    html = _env.get_template("merged.html.j2").render(title=title, items=items, data_json=safe_json(embedded))
    return html, skipped
