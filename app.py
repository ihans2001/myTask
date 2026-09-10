# app.py
# -------------------------------------------------------------
# AX SW 관리 시스템 - 메인 진입점
# st.navigation()/st.Page()로 4개 페이지를 명시적으로 구성한다.
#   1) Greeting        - 자기소개 (introduction.py)
#   2) Automated tasks - 자동화 태스크 목록 (automation_tasks_page.py)
#   3) SaaS Demos      - 5개 핵심 모듈 탭 데모 (saas_demos_page.py)
#   4) Q&A             - 문의처 및 FAQ (qna_page.py)
# -------------------------------------------------------------

import streamlit as st

import introduction
import automation_tasks_page
import saas_demos_page
import qna_page  # <--- Q&A 페이지 모듈 추가

st.set_page_config(page_title="AX SW 관리 시스템", layout="wide")

pg = st.navigation([
    st.Page(introduction.main, title="Greeting", icon="👋", url_path="greeting", default=True),
    st.Page(automation_tasks_page.render, title="Automated tasks", icon="🗂️", url_path="automated-tasks"),
    st.Page(saas_demos_page.render, title="SaaS Live Demos", icon="⚡", url_path="saas-demos"),
    st.Page(qna_page.render, title="Q&A", icon="❓", url_path="qna"),  # <--- 신규 메뉴 추가
])

pg.run()
