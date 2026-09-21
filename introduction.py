import streamlit as st

def main():
    st.title("안녕하세요!")

    #st.header("💡 다짐")
    st.write("""

    자동차, 방위, 우주 산업 분야에서 축적한 경험과 전문성을 바탕으로,
    자동차 SW 개발 프로젝트의 성공적인 수행을 위한 엔지니어링 용역 업무를 성실히 수행할 준비가 되어 있습니다.
    국제 표준인 Automotive SPICE 4.0 Provisional Assessor와 ISO 21434 Lead Auditor 자격을 보유하여
    전문성을 더욱 강화하였으며,
    새로운 도전 앞에서도 끊임없이 배우고 발전하며,
    귀사와 함께 가치 있는 성과를 만들어가겠습니다.
    
    """)

    st.header("👤 소개")

    st.subheader("🎓 학력")
    st.write("""
    - 서울대학교 기계공학 학사
    - 서울대학교 기계공학 석사 (섬유 강화 복합재료)
    - 서울대학교 기계항공공학 박사 (로켓추진, 메탄엔진)
    """)

    st.subheader("💼 경력")
    st.write("""
    - 현대우주항공('94.1~): 액체로켓 엔진 개발, 상업용 가스터빈 부품 개발
    - 현대모비스: 액체로켓 엔진 개발
    - 현대로템: 유도무기 엔진 개발, 연구기획, 신규사업 기획
    - 현대오트론: 경영기획, ECU ASW 로직 개발
    - 현대케피코(~'26.7): ECU ASW 로직 개발, 프로젝트 관리/평가, 안전분석, 사이버보안 TARA/Concept Design/취약점모니터링/성적서 게이트 리뷰
    """)

    st.subheader("📜 자격")
    st.write("""
    - Automotive SPICE 4.0 Provisional Assessor
    - ISO 21434 Lead Auditor
    """)

    st.header("📝 논문 및 학술 활동")

    st.subheader("🚗 자동차 분야") 
    st.write("""
    - SAE International Journal of Engines: 3편 (실화 연구, ESCI, 2020~2022)
    - International Journal of Automotive Technology: 1편 (실화 연구, 2022)
    - SAE WCX2025 (디트로이트): 1편 (아키텍처링 + SW FMEA 연계 수행, 2025)
    - KSAE 기술논문집: 2편 (실화 연구, 2022, 2023)
    - KSAE 춘계/추계 학술대회: 4편 (New SW FMEA, 취약점 모니터링, CS 성적서 게이트 리뷰, 2024~2026)
    - KSAE 춘계/추계 학술대회: 2편 예정 (SR 지향 추적성 연결, 시퀀스 다이어그램과 사이버보안 공격표면 연결, 2026)
    - 한국정보보호학회 하계학술대회: 1편 (TARA New AFR 방법론, 2026)
    - 현대차 그룹 학술대회: 3편 (실화 연구, SW FMEA + FTA/DFA)
    """)

    st.subheader("🚀 항공우주 분야")    
    st.write("""
    - JPP, 추진공학회, 항공우주학회 학회지/학술대회 45편 발표 (구글학술에서 'Poonggyoo Han' 또는 'Poong Gyoo Han'으로 검색해 보세요)
    - 추진공학회 기술상 수상 (2009)
    - 추진공학회 우수 논문 발표상 수상 (2005)
    """)

    st.subheader("🧾 특허")
    st.write("""
    - 원천특허 4건 (Engine Control Unit, 투스 타임 기반 최초 특허)
    - 특허 16건 (Engine Control Unit, 원천특허 회피 방어)
    - 기타 다수
    """)

    st.header("🤝 연락")
    st.subheader("📧 메일")
    st.write("""
    - ihans2001@gmail.com
    - ihans2001@naver.com
    - pghan@hanmail.net
    """)   

    #st.subheader("📞 전화번호")
    #st.write("""
    #- 010-9268-5335
    #""")       

if __name__ == "__main__":
    main()
