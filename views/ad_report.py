"""쿠팡 광고 보고서 생성 페이지 (app.py 라우터에서 실행)."""
from __future__ import annotations

import calendar
import hashlib
import io
import json
from datetime import date

import pandas as pd
import streamlit as st

from core.analysis import Settings, build_context, choose_granularity, make_buckets
from core.data import ReportFileError, campaign_table, date_span, guess_meta, load_report
from core.render import extract_report_data, extract_title, merge_reports, render_report

if "reports" not in st.session_state:
    st.session_state.reports = {}  # {탭 이름: html}


# 업로드 데이터는 서버 메모리에만 잠시 캐시 (1시간, 최대 8개)
@st.cache_data(show_spinner="엑셀을 읽는 중…", ttl=3600, max_entries=8)
def _load(data: bytes, name: str, window: str) -> pd.DataFrame:
    return load_report(io.BytesIO(data), name, window)


def _days_in(month: int | None) -> int:
    if not month:
        return 0
    today = date.today()
    year = today.year if month <= today.month else today.year - 1
    return calendar.monthrange(year, month)[1]


def _num(v) -> float | None:
    try:
        f = float(v)
        return f if f > 0 else None
    except (TypeError, ValueError):
        return None


st.title("📊 쿠팡 광고 보고서 생성기")
st.caption("쿠팡 Wing › 광고 보고서 엑셀(직전월 · 금월)을 올리면 계정 1개 단위 HTML 보고서를 만들고, 여러 보고서를 한 파일로 합칠 수 있습니다.")

tab_make, tab_merge = st.tabs(["📄 보고서 생성", "🧩 보고서 통합"])

# =====================================================================
# 탭 1. 계정 보고서 생성
# =====================================================================
def make_tab():
    c1, c2 = st.columns(2)
    prev_file = c1.file_uploader("① 직전월 광고 보고서", type=["xlsx", "xlsm", "xls", "csv"], key="prev_file")
    cur_file = c2.file_uploader("② 금월 광고 보고서", type=["xlsx", "xlsm", "xls", "csv"], key="cur_file")

    if not (prev_file and cur_file):
        st.info("직전월과 금월 광고 보고서 파일을 모두 업로드하세요. (예: `브랜드명_8월 광고보고서.xlsx`, `브랜드명_9월 광고보고서.xlsx`)")
        return
    pm, cm = guess_meta(prev_file.name), guess_meta(cur_file.name)
    if pm["account"] and cm["account"] and pm["account"] != cm["account"]:
        st.warning(f"두 파일의 계정명이 다릅니다: '{pm['account']}' vs '{cm['account']}'. 같은 계정 파일인지 확인하세요.")

    # ---------------- 기본 정보 ----------------
    st.subheader("기본 정보")
    a1, a2, a3 = st.columns([2, 1, 1])
    account = a1.text_input("계정(브랜드)명", value=cm["account"] or pm["account"] or "")
    written = a2.text_input("작성일", value=date.today().strftime("%Y.%m.%d"))
    window = a3.radio("광고 전환 기준", ["14일", "1일"], horizontal=True,
                      help="쿠팡 기본 화면의 광고전환매출은 14일 기준입니다.")

    # ---------------- 파일 로드 ----------------
    try:
        prev_df = _load(prev_file.getvalue(), prev_file.name, window)
        cur_df = _load(cur_file.getvalue(), cur_file.name, window)
    except ReportFileError as e:
        st.error(str(e))
        return
    prev_span, cur_span = date_span(prev_df), date_span(cur_df)

    def _span_days(span, meta):
        if span:  # 파일 안 날짜가 가장 정확
            return (span[1] - span[0]).days + 1
        return meta["days"] or _days_in(meta["month"])  # 파일명 기간(예: 9월16~9월22일) → 월 일수

    def _label(meta, fallback):
        return meta["label"] or (f"{meta['month']}월" if meta["month"] else fallback)

    b1, b2, b3, b4 = st.columns(4)
    prev_label = b1.text_input("직전월 표기", value=_label(pm, "직전월"))
    prev_days = b2.number_input("직전월 집계 일수", min_value=0, max_value=366, value=_span_days(prev_span, pm),
                                help="일평균 계산용. 일별 파일이면 날짜에서 자동 계산됩니다. 0이면 일평균을 표시하지 않습니다.")
    cur_label = b3.text_input("금월 표기", value=_label(cm, "금월"))
    cur_days = b4.number_input("금월 집계 일수", min_value=0, max_value=366, value=_span_days(cur_span, cm),
                               help="월 중간 보고(예: 1~14일)라면 실제 일수로 바꿔주세요.")

    if prev_label.strip() == cur_label.strip():
        st.warning(f"직전월과 금월 표기가 모두 '{cur_label}'입니다. 보고서에 '{cur_label} vs {cur_label}'로 표시되니 "
                   "직전월 표기를 확인해 주세요. (파일명에서 월을 자동으로 읽습니다)")

    # ---------------- 총매출 (월 합계 또는 구간별) ----------------
    gran = choose_granularity(prev_df, cur_df)
    bucket_sales: dict = {}
    if gran:
        unit = "1일" if gran == "day" else "7일"
        st.markdown(f"**총매출 입력 · {unit} 단위** — 일별 보고서가 감지되어 추이 차트를 {unit} 단위로 그립니다. "
                    "한 칸이라도 비어 있으면 차트는 생성되지 않습니다.")
        saved_sales = st.session_state.setdefault("bucket_sales", {})
        cols = st.columns(2)
        totals = {}
        for col, side, df, label in ((cols[0], "prev", prev_df, prev_label), (cols[1], "cur", cur_df, cur_label)):
            buckets = make_buckets(df, side, gran)
            table = pd.DataFrame({
                "key": [b["key"] for b in buckets],
                "구간": [b["label"] for b in buckets],
                "광고비(참고)": [round(b["cost"]) for b in buckets],
                "총매출(원)": pd.array([saved_sales.get(b["key"]) for b in buckets], dtype="Float64"),  # 빈칸은 공란 표시
            })
            col.caption(f"{label} ({buckets[0]['start']:%m/%d}~{buckets[-1]['end']:%m/%d})")
            edited_sales = col.data_editor(
                table, hide_index=True, width="stretch", disabled=["key", "구간", "광고비(참고)"],
                column_order=["구간", "광고비(참고)", "총매출(원)"],
                column_config={"광고비(참고)": st.column_config.NumberColumn(format="%d"),
                               "총매출(원)": st.column_config.NumberColumn(min_value=0, step=10000, format="%d")},
                key=f"sales::{side}::{gran}::{prev_file.name if side == 'prev' else cur_file.name}",
            )
            vals = {r["key"]: (None if pd.isna(r["총매출(원)"]) else _num(r["총매출(원)"]))
                    for _, r in edited_sales.iterrows()}
            saved_sales.update(vals)
            bucket_sales.update({k: v for k, v in vals.items() if v})
            totals[side] = sum(vals.values()) if all(vals.values()) else 0.0
        prev_sales, cur_sales = totals["prev"], totals["cur"]
        st.caption("핵심 지표의 총매출은 구간별 입력값의 합계로 계산됩니다.")
    else:
        missing = [lbl for lbl, sp in ((prev_label, prev_span), (cur_label, cur_span)) if not sp]
        st.info(f"📅 {', '.join(missing)} 파일에 날짜 칸이 없어(기간 합계 파일) 추이 차트를 **기간 합계 2점**으로 그립니다. "
                "꺾은선으로 보려면 쿠팡 Wing 광고 보고서를 **일별**로 내려받아 두 파일 모두 다시 올려주세요. "
                "(월 파일 → 7일 단위, 7·14일 파일 → 1일 단위)")
        s1, s2 = st.columns(2)
        prev_sales = s1.number_input(f"{prev_label} 총매출(원)", min_value=0, step=100000, value=0,
                                     help="광고 보고서 엑셀에는 총매출(자연매출 포함)이 없어 직접 입력합니다. 0이면 '입력 필요'로 표시됩니다.")
        cur_sales = s2.number_input(f"{cur_label} 총매출(원)", min_value=0, step=100000, value=0)

    with st.expander("⚙️ 분석 기준 (고급)"):
        g1, g2, g3, g4 = st.columns(4)
        min_cost = g1.number_input("③ 최고효율 후보 최소 광고비(원)", min_value=0, step=10000, value=50000,
                                   help="광고비가 너무 적은 캠페인의 ROAS 착시를 막습니다.")
        min_product_cost = g2.number_input("⑤ ROAS 순위 상품 최소 광고비(원)", min_value=0, step=1000, value=10000)
        min_clicks = g3.number_input("⑤ CVR 순위 상품 최소 클릭수", min_value=1, step=1, value=10)
        top_n = g4.number_input("TOP N", min_value=3, max_value=10, value=5)

    # ---------------- 캠페인 설정 (목표 ROAS / 일 예산) ----------------
    st.subheader("캠페인 설정 · 목표 ROAS / 일 예산")
    st.caption("광고 보고서에는 목표 ROAS·일 예산이 없어 광고센터 값을 입력합니다. 비워두면 보고서에 '확인 필요'로 표시됩니다. "
               "입력값은 JSON으로 저장해 다음 달에 불러올 수 있습니다.")
    camps = campaign_table(cur_df)
    camps = camps[camps["cost"] > 0]
    settings_key = f"targets::{account}"
    saved: dict = st.session_state.get(settings_key, {})

    up = st.file_uploader("저장한 캠페인 설정 불러오기 (.json)", type=["json"], key="targets_json")
    if up is not None:
        try:
            saved = {str(k): v for k, v in json.loads(up.getvalue().decode("utf-8")).items()}
            st.session_state[settings_key] = saved
        except (ValueError, AttributeError):
            st.error("캠페인 설정 JSON 형식이 올바르지 않습니다.")

    editor_df = pd.DataFrame({
        "캠페인 ID": camps["id"].astype(str),
        "캠페인명": camps["name"],
        "금월 광고비": camps["cost"].round().astype(int),
        "목표 ROAS(%)": [_num((saved.get(i) or {}).get("target_roas")) for i in camps["id"].astype(str)],
        "일 예산(원)": [_num((saved.get(i) or {}).get("daily_budget")) for i in camps["id"].astype(str)],
    })
    edited = st.data_editor(
        editor_df, hide_index=True, width="stretch",
        disabled=["캠페인 ID", "캠페인명", "금월 광고비"],
        column_config={
            "금월 광고비": st.column_config.NumberColumn(format="%d"),
            "목표 ROAS(%)": st.column_config.NumberColumn(min_value=0, step=5, format="%d"),
            "일 예산(원)": st.column_config.NumberColumn(min_value=0, step=1000, format="%d"),
        },
        key=f"editor::{cur_file.name}",
    )
    targets = {
        str(r["캠페인 ID"]): {"target_roas": _num(r["목표 ROAS(%)"]), "daily_budget": _num(r["일 예산(원)"])}
        for _, r in edited.iterrows()
    }
    st.session_state[settings_key] = {**saved, **targets}
    st.download_button("💾 캠페인 설정 저장 (.json)",
                       data=json.dumps(st.session_state[settings_key], ensure_ascii=False, indent=2),
                       file_name=f"{account or '계정'}_캠페인설정.json", mime="application/json")

    # ---------------- 분석 ----------------
    settings = Settings(
        account=account or "계정", prev_label=prev_label, cur_label=cur_label,
        prev_days=int(prev_days), cur_days=int(cur_days),
        prev_total_sales=float(prev_sales), cur_total_sales=float(cur_sales),
        window=window, min_cost=float(min_cost), min_product_cost=float(min_product_cost),
        min_clicks=int(min_clicks), top_n=int(top_n), written=written, targets=targets,
        bucket_sales=bucket_sales,
    )
    ctx = build_context(prev_df, cur_df, settings)

    st.subheader("미리 보는 핵심 수치")
    m = st.columns(5)
    for col, k in zip(m, ctx["kpis"]):
        col.metric(k["label"], k["cur"] or "입력 필요",
                   k["delta"]["text"] if k["delta"] else None, delta_color="off")

    # ---------------- ④ 3단계 요약 (수정 가능) ----------------
    st.subheader("④ 개선점 & 이슈 요약 (자동 초안 · 수정 가능)")
    st.caption("한 줄에 한 가지만 짧게 적어주세요. `•` 줄은 항목, `→` 줄은 결론 박스로 표시되고 "
               "`▲ 12%`는 빨간색, `▼ 3.5%p`는 파란색으로 자동 강조됩니다.")
    sig = hashlib.md5(json.dumps([s["text"] for s in ctx["steps"]], ensure_ascii=False).encode()).hexdigest()
    if st.session_state.get("steps_sig") != sig:  # 데이터·기준이 바뀌면 초안을 새로 채움
        st.session_state.steps_sig = sig
        for i, stp in enumerate(ctx["steps"]):
            st.session_state[f"step_title_{i}"] = stp["title"]
            st.session_state[f"step_text_{i}"] = stp["text"]
    cols = st.columns(3)
    for i, col in enumerate(cols):
        col.text_input(f"STEP {i + 1} 제목", key=f"step_title_{i}")
        col.text_area(f"STEP {i + 1} 내용", key=f"step_text_{i}", height=140)
    ctx["steps"] = [{"title": st.session_state[f"step_title_{i}"], "text": st.session_state[f"step_text_{i}"]}
                    for i in range(3)]

    # ---------------- 생성 ----------------
    st.divider()
    html = render_report(ctx)
    fname = f"쿠팡_광고_보고서_{settings.account}_{cur_label}.html"
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("⬇️ HTML 보고서 다운로드", data=html.encode("utf-8"), file_name=fname,
                       mime="text/html", type="primary", width="stretch")
    tab_label = f"{settings.account} 쿠팡 보고서"
    if d2.button("➕ 통합 목록에 추가", width="stretch",
                 help="'보고서 통합' 탭에서 다른 계정 보고서와 합칠 수 있습니다."):
        st.session_state.reports[tab_label] = html
        st.success(f"'{tab_label}'을(를) 통합 목록에 추가했습니다. (현재 {len(st.session_state.reports)}개)")

    st.subheader("보고서 미리보기")
    st.iframe(html, height=900)

with tab_make:
    make_tab()

# =====================================================================
# 탭 2. 보고서 통합
# =====================================================================
with tab_merge:
    st.markdown("생성한 계정별 보고서를 **상단 탭으로 전환되는 하나의 HTML 파일**로 합칩니다.")
    uploads = st.file_uploader("기존에 저장한 보고서 HTML 추가 (여러 개 선택 가능)", type=["html", "htm"],
                               accept_multiple_files=True, key="merge_uploads")

    pool: dict[str, str] = dict(st.session_state.reports)
    for f in uploads or []:
        text = f.getvalue().decode("utf-8", errors="replace")
        if 'id="embedded-reports"' in text:
            st.warning(f"'{f.name}'은(는) 이미 통합된 보고서라 제외했습니다. 계정별 보고서를 올려주세요.")
            continue
        data = extract_report_data(text)
        label = f"{data['account']} 쿠팡 보고서" if data else extract_title(text, f.name)
        base, n = label, 2
        while label in pool and pool[label] != text:
            label, n = f"{base} ({n})", n + 1
        pool[label] = text

    if not pool:
        st.info("'보고서 생성' 탭에서 **➕ 통합 목록에 추가**를 누르거나, 저장해 둔 보고서 HTML을 업로드하세요.")
    else:
        if st.session_state.reports and st.button("🗑️ 생성 탭에서 추가한 목록 비우기"):
            st.session_state.reports = {}
            st.rerun()

        order = st.multiselect("통합할 보고서 (선택한 순서 = 탭 순서)", list(pool), default=list(pool))
        if order:
            st.markdown("**탭 이름**")
            labels = []
            cols = st.columns(min(len(order), 4))
            for i, name in enumerate(order):
                labels.append(cols[i % len(cols)].text_input(f"탭 {i + 1}", value=name, key=f"tablabel::{name}"))

            t1, t2 = st.columns([2, 1])
            first = extract_report_data(pool[order[0]]) or {}
            title = t1.text_input("통합 보고서 제목", value=f"쿠팡 광고 보고서 · {first.get('cur_label') or f'{date.today():%Y년 %m월}'}")
            add_summary = t2.checkbox("계정 통합 요약 탭 추가", value=len(order) >= 2,
                                      help="각 보고서에 내장된 수치로 전 계정 합산·비교 탭을 만듭니다. (2개 이상일 때)")

            merged, skipped = merge_reports([{"label": l, "html": pool[n]} for l, n in zip(labels, order)],
                                            title=title, add_summary=add_summary)
            if add_summary and skipped:
                st.warning("이 앱으로 만들지 않은 보고서는 통합 요약에서 제외됩니다: " + ", ".join(skipped))
            st.download_button("⬇️ 통합 보고서 다운로드", data=merged.encode("utf-8"),
                               file_name=f"쿠팡_광고_보고서_통합_{date.today():%Y%m}.html",
                               mime="text/html", type="primary")
            st.subheader("통합 보고서 미리보기")
            st.iframe(merged, height=900)
