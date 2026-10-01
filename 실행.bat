@echo off
chcp 65001 > nul
cd /d "%~dp0"
set COUPANG_REPORT_LOCAL=1
py -m pip install -q -r requirements.txt
py -m streamlit run app.py
pause
