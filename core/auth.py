"""접속 비밀번호 잠금 (모든 페이지 공용)."""
from __future__ import annotations

import hmac
import os

import streamlit as st


def _configured_password() -> str | None:
    """Streamlit Secrets의 APP_PASSWORD. 설정 파일이 없으면 None."""
    try:
        value = st.secrets.get("APP_PASSWORD")
    except Exception:  # secrets.toml 자체가 없을 때
        return None
    return str(value) if value else None


def require_login(title: str = "🔒 쿠팡 보고서 생성기") -> None:
    """로컬 실행(실행.bat)은 건너뛰고, 온라인에서 비밀번호가 없으면 열지 않는다."""
    if os.environ.get("COUPANG_REPORT_LOCAL") == "1" or st.session_state.get("authed"):
        return
    expected = _configured_password()
    st.title(title)
    if expected is None:
        st.error("접속 비밀번호가 아직 설정되지 않았습니다. 관리자가 Streamlit 앱 설정 › Secrets에 "
                 "`APP_PASSWORD`를 등록해야 사용할 수 있습니다.")
        st.stop()
    with st.form("login"):
        pw = st.text_input("접속 비밀번호", type="password")
        ok = st.form_submit_button("입장", type="primary")
    if ok:
        if hmac.compare_digest(pw.encode(), expected.encode()):
            st.session_state.authed = True
            st.rerun()
        st.error("비밀번호가 올바르지 않습니다.")
    st.stop()
