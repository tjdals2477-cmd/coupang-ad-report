"""Jinja2로 계정 보고서 HTML을 만들고, 여러 보고서를 탭형 단일 HTML로 병합한다."""
from __future__ import annotations

import json
import re
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup, escape

from .analysis import delta, pct, won

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"
_env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=select_autoescape(["html", "j2"]))

_DELTA_RE = re.compile(r"([▲▼])\s?([\d.,]+\s?%p?)")


def colorize_deltas(text: str) -> Markup:
    """이스케이프 후 '▲ 12.3%'는 빨강, '▼ 4.5%p'는 파랑 span으로 감싼다."""
    def wrap(m):
        cls = "up" if m.group(1) == "▲" else "down"
        return f'<span class="d-{cls}">{m.group(1)} {m.group(2)}</span>'
    return Markup(_DELTA_RE.sub(wrap, str(escape(text))))


def step_lines(text: str) -> list[dict]:
    """요약 텍스트를 줄 단위로 나눈다: '→'로 시작하면 결론, 나머지는 불릿."""
    out = []
    for line in str(text).splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("→"):
            out.append({"kind": "conc", "html": colorize_deltas(line.lstrip("→ ").strip())})
        else:
            out.append({"kind": "item", "html": colorize_deltas(line.lstrip("•-· ").strip())})
    return out


_env.filters["step_lines"] = step_lines

DATA_TAG_RE = re.compile(r'<script id="report-data" type="application/json">(.*?)</script>', re.S)
TITLE_RE = re.compile(r"<title>(.*?)</title>", re.S)


def safe_json(obj) -> str:
    """<script> 안에 넣어도 안전한 JSON (</script> 조기 종료 방지)."""
    return (json.dumps(obj, ensure_ascii=False)
            .replace("<", "\\u003C").replace(">", "\\u003E").replace("&", "\\u0026"))


def report_data(ctx: dict) -> dict:
    """병합 시 '계정 통합 요약' 탭을 만들 수 있도록 보고서에 내장하는 요약 수치."""
    s = ctx["s"]
    return {
        "version": 1,
        "account": s.account,
        "prev_label": s.prev_label, "cur_label": s.cur_label,
        "prev": {**ctx["prev"], "total_sales": s.prev_total_sales},
        "cur": {**ctx["cur"], "total_sales": s.cur_total_sales},
        "n_campaigns": ctx["n_campaigns"],
        "issue": ctx["issue"]["cur"]["name"] if ctx["issue"] else None,
        "best": ctx["best"]["cur"]["name"] if ctx["best"] else None,
        "best_roas": ctx["best"]["cur"]["roas"] if ctx["best"] else None,
    }


def render_report(ctx: dict) -> str:
    tpl = _env.get_template("report.html.j2")
    trend = ctx.get("trend")
    trend_json = safe_json({
        "series": [{"label": sr["label"], "color": sr["color"]} for sr in trend["series"]],
        "cols": trend["tooltip"],
    }) if trend else ""
    return tpl.render(**ctx, data_json=safe_json(report_data(ctx)), trend_json=trend_json)


def extract_report_data(html: str) -> dict | None:
    m = DATA_TAG_RE.search(html)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return None


def extract_title(html: str, fallback: str) -> str:
    m = TITLE_RE.search(html)
    return m.group(1).strip() if m else fallback


def render_summary(datas: list[dict], title: str) -> str:
    """보고서들에 내장된 report-data로 계정 통합 요약 페이지를 만든다."""
    def tot(side, key):
        return sum(d[side].get(key, 0) or 0 for d in datas)

    total = {}
    for side in ("prev", "cur"):
        cost, rev = tot(side, "cost"), tot(side, "revenue")
        clicks, orders = tot(side, "clicks"), tot(side, "orders")
        sales_known = all(d[side].get("total_sales") for d in datas)
        total[side] = {"cost": cost, "revenue": rev,
                       "roas": rev / cost * 100 if cost else 0,
                       "cvr": orders / clicks * 100 if clicks else 0,
                       "total_sales": tot(side, "total_sales") if sales_known else 0}
    rows = []
    for d in datas:
        p, c = d["prev"], d["cur"]
        rows.append({
            "account": d["account"], "prev_label": d["prev_label"], "cur_label": d["cur_label"],
            "cost": won(c["cost"]), "cost_d": delta(c["cost"], p["cost"]),
            "revenue": won(c["revenue"]), "revenue_d": delta(c["revenue"], p["revenue"]),
            "roas": pct(c["roas"]), "roas_d": delta(c["roas"], p["roas"], "pp"),
            "cvr": pct(c["cvr"]), "cvr_d": delta(c["cvr"], p["cvr"], "pp"),
            "sales": won(c["total_sales"]) if c.get("total_sales") else "입력 필요",
            "ad_ratio": pct(c["cost"] / c["total_sales"] * 100, 1) if c.get("total_sales") else "-",
            "best": d.get("best") or "-", "issue": d.get("issue") or "-",
            "share": c["cost"] / total["cur"]["cost"] * 100 if total["cur"]["cost"] else 0,
        })
    labels = {d["cur_label"] for d in datas}
    kpis = []
    for label, key, fmt, kind in [("총 광고집행비용", "cost", won, "pct"), ("총매출", "total_sales", won, "pct"),
                                  ("광고 전환 매출", "revenue", won, "pct"), ("ROAS", "roas", pct, "pp"),
                                  ("전환율(CVR)", "cvr", pct, "pp")]:
        pv, cv = total["prev"][key], total["cur"][key]
        known = key != "total_sales" or (pv and cv)
        kpis.append({"label": label, "prev": fmt(pv) if known else None, "cur": fmt(cv) if known else None,
                     "delta": delta(cv, pv, kind) if known else None})
    tpl = _env.get_template("summary.html.j2")
    return tpl.render(title=title, rows=rows, kpis=kpis,
                      prev_label=datas[0]["prev_label"], cur_label=datas[0]["cur_label"],
                      mixed_period=len(labels) > 1)


def merge_reports(reports: list[dict], title: str = "쿠팡 광고 보고서",
                  add_summary: bool = True) -> tuple[str, list[str]]:
    """reports: [{"label": 탭 이름, "html": 보고서 HTML}, ...] (순서 = 탭 순서).
    반환: (병합 HTML, 요약 데이터가 없어 통합 요약에서 빠진 보고서 이름 목록)"""
    if not reports:
        raise ValueError("병합할 보고서가 없습니다.")
    items = [{"id": f"r{i}", "label": r["label"], "html": r["html"]} for i, r in enumerate(reports)]
    skipped = []
    if add_summary:
        datas = []
        for r in reports:
            d = extract_report_data(r["html"])
            (datas.append(d) if d else skipped.append(r["label"]))
        if len(datas) >= 2:
            items.append({"id": "summary", "label": "계정 통합 보고서",
                          "html": render_summary(datas, f"{title} · 계정 통합")})
    embedded = {it["id"]: {"label": it["label"], "html": it["html"]} for it in items}
    tpl = _env.get_template("merged.html.j2")
    html = tpl.render(title=title, items=items, data_json=safe_json(embedded))
    return html, skipped
