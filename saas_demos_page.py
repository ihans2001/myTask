# saas_demos_page.py
# -------------------------------------------------------------
# "SaaS Demos" 페이지: AX SW 관리 시스템의 실제 기능 5종을 탭으로 묶어 보여준다.
# 각 탭은 목업이 아니라 이미 구현되어 검증된 실제 모듈을 그대로 호출한다.
# -------------------------------------------------------------

import streamlit as st

import kpms_schedule_page
import evms_dashboard_page
import mail_tone_page
import tara_afr_page
import hash_verification_page
import dfa_fta_page
import ssot_release_page
import mitre_monitoring_page
import cvss_calculator_page


def render():
    st.title("⚡ AX SW 개발업무 자동화 - SaaS Live Demo")
    st.caption("검증된 9가지 핵심 모듈의 웹 기반 실행 파이프라인")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "📅 테일러링/일정 자동생성",
        "📊 EVMS 성과 관리 대시보드",
        "✉️ 메일 톤&매너 및 리스크 분석",
        "🌳 TARA AFR/공격트리",
        "🔐 SW 산출물 해시 검증",
        "🩻 DFA/FTA 고장 다이어그램",
        "📦 SSoT 릴리즈 준비도/노트작성",
        "🛰️ MITRE CVE 모니터링",
        "🧮 CVSS v3.1/v4.0 계산기",
    ])

    with tab1:
        kpms_schedule_page.render()
    with tab2:
        evms_dashboard_page.render()
    with tab3:
        mail_tone_page.render()
    with tab4:
        tara_afr_page.render()
    with tab5:
        hash_verification_page.render()
    with tab6:
        dfa_fta_page.render()
    with tab7:
        ssot_release_page.render()
    with tab8:
        mitre_monitoring_page.render()
    with tab9:
        cvss_calculator_page.render()
