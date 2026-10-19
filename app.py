"""쿠팡 보고서 생성기 (Streamlit) — 페이지 라우터.

실행:  streamlit run app.py
페이지: views/ad_report.py(광고 보고서) · views/monthly_report.py(월간 성과 보고서)
"""
import streamlit as st

from core.auth import require_login

st.set_page_config(page_title="쿠팡 보고서 생성기", page_icon="📊", layout="wide")
st.markdown("""
<style>
  .block-container{padding-top:2rem; max-width:1200px;}
  h1{font-size:1.9rem !important;}
  div[data-testid="stMetricValue"]{font-size:1.2rem;}
</style>
""", unsafe_allow_html=True)

require_login()

st.navigation([
    st.Page("views/ad_report.py", title="광고 보고서", icon="📊", default=True),
    st.Page("views/monthly_report.py", title="월간 성과 보고서", icon="📈"),
]).run()
