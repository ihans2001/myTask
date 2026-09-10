# qna_page.py
import streamlit as st

def render():
    st.title("❓ Q&A")
    st.caption("시스템 이용 및 기술 문의에 대해 안내해 드립니다.")
    st.markdown("---")

    # 1. 이메일 문의 안내 영역
    st.subheader("📩 문의 안내")
    st.info("""
    궁금하신 점이나 추가 기능 요청, 기술 지원이 필요하신 경우, 아래 이메일 주소 중 하나를 선택, 문의해 주시기 바랍니다.  
    **질문을 남겨주시면, 24시간 이내에 회신해 드리겠습니다.**
    """)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("📧 **Primary Gmail:** [ihans2001@gmail.com](mailto:ihans2001@gmail.com)")
    with col2:
        st.markdown("📧 **Naver Mail:** [ihans2001@naver.com](mailto:ihans2001@naver.com)")
    with col3:
        st.markdown("📧 **Hanmail:** [pghan@hanmail.net](mailto:pghan@hanmail.net)")

    st.markdown("---")

    # 2. 자주 묻는 질문 (Q&A 목록) 영역
    st.subheader("💡 자주 묻는 질문 (FAQ)")
    st.write("메일로 문의하시기 전에 아래에서 주요 질의응답 내용을 확인해 보세요.")

    with st.expander("Q1. SaaS 데모를 하면서, 시스템 보안 및 데이터 유출 위험은 없나요?"):
        st.write("""
        업로드하신 모든 엑셀 및 메일 파일은 외부 웹 서버에 저장되지 않고 로컬 휘발성 메모리 세션 내에서만 즉시 처리되므로 
        보안 및 데이터 유출 걱정 없이 안심하고 사용하실 수 있습니다.
        """)

    with st.expander("Q2. 데모와 실제 기능과는 어떤 차이가 있나요?"):
        st.write("""
        자동화 태스크 34개 항목은 VBA, Python 및 LLM과 연계하여 동작합니다. 현재의 데모는 일부 태스크에 국한해서, 일부 기능을 맛보기 목적으로 재구성한 것입니다.
        실제 태스크에는 다양한 현장 환경에서 사용가능한 기능이 탑재되어 있으며, 선택 태스크에 대한 과제를 수행하게 되면, 여러분의 업무 환경을 분석하고, 분석 결과를 바탕으로 모든 함수는 커스터마이징되어야 합니다. 
        """)

    with st.expander("Q3. 현재의 데모 이외에, 추가 데모를 확인할 수 있나요?"):
        st.write("""
        메일로 문의를 남겨주세요! 구현 가능한 데모는 추가할 수 있습니다. 
        데모 추가 외에도, 세부 검토를 위한 가이드 자료를 별도로 제공해 드릴 수 있습니다.
        """)        

    with st.expander("Q4. macOS 환경에서도 데모가 정상 동작하나요?"):
        st.write("""
        네, 현재의 데모는 OpenPyXL, Pandas, Hashlib 등 순수 Python 기반 알고리즘으로 구축되어 있어, 
        Windows뿐만 아니라 Mac Studio(macOS) 환경에서도 100% 동일하고 안정적으로 동작합니다. 
        단, TARA 공격트리 렌더링을 위해 Mac 서버에 `brew install graphviz` 바이너리 설치가 완료되어 있는지 확인해 주세요.
        """)        

    with st.expander("Q5. 태스크 내용이 쉽게 이해가 되지 않고, 데모 사용도 어렵습니다!"):
        st.write("""
        태스크를 핵심기능과 인포그래픽으로 간략하게 구성하였으므로, 태스크에 대한 이해가 쉽지 않을 수 있습니다. 
        한편, 데모에서는 일부 기능을 직관적으로 이해하실 수 있도록 구성했습니다만, 처음 접하실 때에는 어려울 수 있습니다.
        메일로 상황을 설명해 주시면, 별도로 안내해 드리겠습니다. 
        """)
