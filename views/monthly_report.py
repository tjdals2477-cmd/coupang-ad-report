"""쿠팡 월간 성과 보고서 페이지: 원본 엑셀 업로드 → 계정별 HTML → 탭형 통합."""
from __future__ import annotations

import hashlib
import re
from datetime import date

import pandas as pd
import streamlit as st

from core.analysis import won
from core.render import extract_report_data, extract_title
from monthly.analysis import MonthlySettings, build_context, resolve_periods
from monthly.data import (AdRaw, InsightData, MonthlyFileError, dashboard_sheets, read_ad_dashboard,
                          read_ad_raw, read_insight)
from monthly.render import merge_monthly, render_monthly

if "monthly_reports" not in st.session_state:
    st.session_state.monthly_reports = {}  # {탭 이름: html}


# 업로드 데이터는 서버 메모리에만 잠시 캐시 (1시간)
@st.cache_data(show_spinner="판매 데이터를 읽는 중…", ttl=3600, max_entries=16)
def _insight(files: tuple) -> InsightData:
    ins = InsightData()
    for name, data in files:
        read_insight(data, name, ins)
    return ins


@st.cache_data(show_spinner="광고 데이터를 읽는 중… (큰 파일은 수십 초)", ttl=3600, max_entries=8)
def _ad_raw(files: tuple) -> AdRaw:
    raw = AdRaw()
    for name, data in files:
        read_ad_raw(data, name, raw)
    return raw


@st.cache_data(ttl=3600, max_entries=8)
def _dash_sheets(data: bytes) -> list[str]:
    return dashboard_sheets(data)


@st.cache_data(ttl=3600, max_entries=8)
def _dash(data: bytes, sheet: str) -> list[dict]:
    return read_ad_dashboard(data, sheet)


def _guess_account(names: list[str]) -> str:
    for n in names:
        head = re.split(r"[_\s]", n)[0]
        if head and not re.match(r"^\d", head):
            return head
    return ""


st.title("📈 쿠팡 월간 성과 보고서")
st.caption("쿠팡 Wing 비즈니스 인사이트(일별·상품별 판매)와 광고 데이터를 올리면 월별 매출·광고비·자연매출·방문자·구매전환율과 "
           "매출 급변 상품을 정리한 HTML 보고서를 만듭니다. 계정별로 만든 뒤 '보고서 통합'에서 한 파일로 합치세요.")

tab_make, tab_merge = st.tabs(["📄 보고서 생성", "🧩 보고서 통합"])


def make_tab():
    st.subheader("① 원본 데이터")
    u1, u2 = st.columns(2)
    ins_files = u1.file_uploader(
        "비즈니스 인사이트 (월별 파일 여러 개)", type=["xlsx"], accept_multiple_files=True, key="m_ins",
        help="Wing › 비즈니스 인사이트에서 월별로 내려받은 파일. '기간별 판매량'(일별)과 '상품별 판매량'(옵션별) 시트를 자동으로 찾습니다. "
             "상품별 판매량만 따로 받은 파일도 함께 올리면 일별 합계와 맞는 쪽을 자동으로 고릅니다.")
    raw_files = u2.file_uploader(
        "광고 raw (선택 · 여러 개)", type=["xlsx", "xlsm"], accept_multiple_files=True, key="m_raw",
        help="광고센터 일자별·옵션별 보고서 또는 광고보고서 xlsm('raw' 시트). 급변 상품의 광고 상태·광고비, 광고 합계 빈 날 보정에 씁니다. "
             "같은 날짜가 여러 파일에 있으면 나중에 올린 파일 기준입니다.")
    if not ins_files:
        st.info("비즈니스 인사이트 파일을 올려주세요. 예: `안트_26년 8월_비즈니스 인사이트.xlsx` … `안트_26년 10월_비즈니스 인사이트.xlsx`")
        return

    with st.expander("광고 주차 대시보드 (선택)", expanded=False):
        dash_file = st.file_uploader("주차별 광고 대시보드 엑셀", type=["xlsx"], key="m_dash",
                                     help="시작일·종료일·광고비용·광고매출 열이 있는 시트. 없으면 광고 raw의 일합계를 씁니다.")

    try:
        ins = _insight(tuple((f.name, f.getvalue()) for f in ins_files))
        raw = _ad_raw(tuple((f.name, f.getvalue()) for f in raw_files)) if raw_files else None
    except MonthlyFileError as e:
        st.error(str(e))
        return

    a1, a2 = st.columns([2, 1])
    account = a1.text_input("계정(브랜드)명", value=_guess_account([f.name for f in ins_files]))
    written = a2.text_input("작성일", value=date.today().strftime("%Y.%m.%d"))

    dash = None
    if dash_file:
        sheets = _dash_sheets(dash_file.getvalue())
        default = sheets.index(account) if account in sheets else 0
        sheet = st.selectbox("대시보드 시트 (계정)", sheets, index=default)
        try:
            dash = _dash(dash_file.getvalue(), sheet)
        except MonthlyFileError as e:
            st.warning(str(e))

    if not ins.daily:
        st.error("'기간별 판매량'(일별) 시트가 있는 파일이 없습니다.")
        return
    auto_cut = max(d for d, v in ins.daily.items() if v["vis"] > 0)
    c1, c2 = st.columns([1, 2])
    cutoff = c1.date_input("집계 마지막 날", value=auto_cut, min_value=min(ins.daily), max_value=max(ins.daily),
                           help="방문자가 있는 마지막 날로 자동 설정됩니다(당일 부분집계 제외).")
    c2.caption(f"판매 데이터 {min(ins.daily):%Y.%m.%d} ~ {max(ins.daily):%Y.%m.%d} · "
               f"상품별 시트 {len(ins.options)}개 · 광고 raw {len(raw.by_day) if raw else 0}일")

    if raw and raw.periods:
        resolve_periods(raw, ins.daily, cutoff)
        st.markdown("**기간 합계 광고 파일** — 날짜 없이 기간 전체로 내려받은 광고보고서. 기간은 파일명에서 읽었으니 확인해 주세요.")
        for i, p in enumerate(raw.periods):
            k1, k2, k3 = st.columns([2.2, 1.4, 1])
            k1.caption(f"📄 {p.source}  \n캠페인 {len(p.campaigns)}개 · 옵션 {len(p.opt)}개 · "
                       f"광고비 {won(p.cost)} · 광고매출 {won(p.rev)}")
            default = (p.s, p.e) if p.s and p.e else (date(cutoff.year, cutoff.month, 1), cutoff)
            rng = k2.date_input("기간", value=default, key=f"m_per::{i}::{p.source}")
            if isinstance(rng, (tuple, list)) and len(rng) == 2:
                p.s, p.e = rng
            p.use_total = k3.checkbox("계정 전체 광고", value=p.use_total, key=f"m_pertot::{i}::{p.source}",
                                      help="모든 캠페인이 들어 있는 파일이면 체크 → 월 광고비·광고매출 합계에도 사용(대시보드보다 우선). "
                                           "캠페인 1개짜리 파일은 체크 해제(옵션별 광고비에만 반영).")

    # ---------------- 보정 · 기준 ----------------
    st.subheader("② 광고 합계 보정 · 분석 기준")
    st.caption("대시보드에도 광고 raw에도 없는 날의 광고비·광고전환매출(14일)을 직접 넣을 수 있습니다. 빈 행은 무시합니다.")
    manual_df = st.data_editor(
        pd.DataFrame({"날짜": pd.Series(dtype="datetime64[ns]"), "광고비": pd.Series(dtype="float"),
                      "광고전환매출": pd.Series(dtype="float")}),
        num_rows="dynamic", hide_index=True, key=f"m_manual::{account}",
        column_config={"날짜": st.column_config.DateColumn(format="YYYY-MM-DD"),
                       "광고비": st.column_config.NumberColumn(min_value=0, format="%d"),
                       "광고전환매출": st.column_config.NumberColumn(min_value=0, format="%d")})
    manual = [(pd.Timestamp(r["날짜"]).date(), float(r["광고비"] or 0), float(r["광고전환매출"] or 0))
              for _, r in manual_df.iterrows() if pd.notna(r["날짜"]) and pd.notna(r["광고비"])]

    with st.expander("⚙️ 급변 상품 · 채널 이동 기준"):
        g1, g2, g3, g4, g5 = st.columns(5)
        min_diff = g1.number_input("최소 증감액(원)", min_value=0, step=100_000, value=500_000)
        min_pct = g2.number_input("최소 증감률(%)", min_value=0, step=5, value=15)
        shift_min = g3.number_input("채널 이동 최소액(원)", min_value=0, step=50_000, value=100_000,
                                    help="상대 채널 옵션이 이 금액 이상 반대로 움직여야 이동으로 봅니다.")
        shift_ratio = g4.number_input("채널 이동 최소 비율(%)", min_value=0, max_value=100, step=5, value=20,
                                      help="내 변동액 대비 상대 채널 변동 비율")
        recent = g5.number_input("광고 상태 판단 일수", min_value=1, max_value=30, value=7)
        pairs_txt = st.text_area("채널 매칭 수동 추가 (한 줄에 `판매자배송 옵션ID, 로켓그로스 옵션ID`)", height=80,
                                 help="옵션명이 달라 자동으로 같은 상품으로 묶이지 않는 경우에만 입력하세요.")
    pairs = [tuple(x.strip() for x in line.split(",")[:2]) for line in pairs_txt.splitlines()
             if line.count(",") >= 1 and all(x.strip().isdigit() for x in line.split(",")[:2])]

    s = MonthlySettings(account=account or "계정", written=written, cutoff=cutoff, min_diff=float(min_diff),
                        min_pct=float(min_pct), shift_min=float(shift_min), shift_ratio=shift_ratio / 100,
                        recent_days=int(recent), manual_ads=manual, channel_pairs=pairs)
    try:
        ctx = build_context(ins, raw, dash, s)
    except ValueError as e:
        st.error(str(e))
        return

    # ---------------- 점검 ----------------
    st.subheader("③ 데이터 점검")
    for level, text in ctx["checks"]:
        (st.warning if level == "warn" else st.caption)(("⚠️ " if level == "warn" else "ℹ️ ") + text)
    if not ctx["checks"]:
        st.caption("✅ 확인할 사항이 없습니다.")
    m = st.columns(len(ctx["kpis"]))
    for col, k in zip(m, ctx["kpis"]):
        col.metric(k["label"], k["rows"][-1]["text"] or "데이터 부족",
                   k["delta"]["text"] if k["delta"] else None, delta_color="off", help=k["how"] or None)

    # ---------------- 배너 문구 ----------------
    st.subheader("④ 요약 문구 (자동 초안 · 수정 가능)")
    sig = hashlib.md5((ctx["auto_headline"] + ctx["auto_headline_sub"]).encode()).hexdigest()
    if st.session_state.get("m_head_sig") != sig:
        st.session_state.m_head_sig = sig
        st.session_state.m_head = ctx["auto_headline"]
        st.session_state.m_head_sub = ctx["auto_headline_sub"]
    ctx["headline"] = st.text_input("핵심 한 줄", key="m_head")
    ctx["headline_sub"] = st.text_input("보조 설명", key="m_head_sub")

    # ---------------- 생성 ----------------
    st.divider()
    html = render_monthly(ctx)
    last = ctx["last"]["short"]
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("⬇️ HTML 보고서 다운로드", data=html.encode("utf-8"),
                       file_name=f"쿠팡_월간_성과_보고서_{s.account}_{last}.html", mime="text/html",
                       type="primary", width="stretch")
    tab_label = f"{s.account} 쿠팡 보고서"
    if d2.button("➕ 통합 목록에 추가", width="stretch"):
        st.session_state.monthly_reports[tab_label] = html
        st.success(f"'{tab_label}'을(를) 통합 목록에 추가했습니다. (현재 {len(st.session_state.monthly_reports)}개)")
    n_mv = sum(len(x["rows"]) for x in ctx["movers"])
    st.caption(f"급변 상품 {n_mv}개 · 매출 {won(ctx['last']['rev'])} ({ctx['last']['label']})")
    st.subheader("보고서 미리보기")
    st.iframe(html, height=900)


with tab_make:
    make_tab()

with tab_merge:
    st.markdown("계정별 월간 보고서를 **상단 탭으로 전환되는 하나의 HTML**로 합치고, 계정 통합 탭을 자동으로 붙입니다.")
    uploads = st.file_uploader("저장해 둔 월간 보고서 HTML 추가 (여러 개)", type=["html", "htm"],
                               accept_multiple_files=True, key="m_merge_uploads")
    pool: dict[str, str] = dict(st.session_state.monthly_reports)
    for f in uploads or []:
        text = f.getvalue().decode("utf-8", errors="replace")
        if 'id="embedded-reports"' in text:
            st.warning(f"'{f.name}'은(는) 이미 통합된 보고서라 제외했습니다.")
            continue
        data = extract_report_data(text)
        label = f"{data['account']} 쿠팡 보고서" if data and data.get("account") else extract_title(text, f.name)
        base, n = label, 2
        while label in pool and pool[label] != text:
            label, n = f"{base} ({n})", n + 1
        pool[label] = text
    if not pool:
        st.info("'보고서 생성' 탭에서 **➕ 통합 목록에 추가**를 누르거나, 저장한 보고서 HTML을 올려주세요.")
    else:
        if st.session_state.monthly_reports and st.button("🗑️ 생성 탭에서 추가한 목록 비우기"):
            st.session_state.monthly_reports = {}
            st.rerun()
        order = st.multiselect("통합할 보고서 (선택한 순서 = 탭 순서)", list(pool), default=list(pool))
        if order:
            cols = st.columns(min(len(order), 4))
            labels = [cols[i % len(cols)].text_input(f"탭 {i + 1}", value=name, key=f"m_tab::{name}")
                      for i, name in enumerate(order)]
            t1, t2 = st.columns([2, 1])
            title = t1.text_input("통합 보고서 제목", value=f"쿠팡 월간 성과 보고서 · {date.today():%Y년 %m월}")
            add_summary = t2.checkbox("계정 통합 탭 추가", value=len(order) >= 2)
            merged, skipped = merge_monthly([{"label": lb, "html": pool[n]} for lb, n in zip(labels, order)],
                                            title=title, add_summary=add_summary)
            if add_summary and skipped:
                st.warning("월간 성과 보고서가 아니어서 통합 탭 합산에서 제외: " + ", ".join(skipped))
            st.download_button("⬇️ 통합 보고서 다운로드", data=merged.encode("utf-8"),
                               file_name=f"쿠팡_월간_성과_보고서_통합_{date.today():%Y%m}.html",
                               mime="text/html", type="primary")
            st.subheader("통합 보고서 미리보기")
            st.iframe(merged, height=900)
