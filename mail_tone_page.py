# mail_tone_page.py
# -------------------------------------------------------------
# 메일 톤&매너 및 커뮤니케이션 리스크 분석. app.py 인라인 코드에서 분리.
# -------------------------------------------------------------

import streamlit as st


def render():
    st.title("✉️ 프로젝트 메일 톤&매너 분석 및 커뮤니케이션 리스크 감지")
    st.caption("mAnal_Tone_mailThread_Using.bas 기반 메일 파일(.msg) 및 텍스트 자동 파싱 모듈")

    #st.info("💡 아웃룩 메일 파일(.msg)을 업로드하거나 메일 본문을 직접 붙여넣으면 텍스트를 파싱하여 보여준 후 리스크를 진단합니다.")
    st.info("💡 아웃룩 메일 파일(.msg)을 업로드하거나 메일 본문을 직접 붙여넣으면 텍스트를 파싱하여 보여준 후 리스크를 진단합니다.\n데모의 진단 범위는 메일 전체이며, 실제 환경에서는 메일 스레드를 분리하여 교신 메일별로 진단하고 이슈 추이까지 모니터링합니다.")


    input_type = st.radio("입력 방식을 선택하세요:", ["📁 아웃룩 메일 파일 (.msg) 업로드", "✍️ 메일 텍스트 직접 붙여넣기"], horizontal=True)

    extracted_text = ""

    if input_type == "📁 아웃룩 메일 파일 (.msg) 업로드":
        msg_file = st.file_uploader("분석할 아웃룩 메일 파일(.msg)을 업로드하세요", type=["msg", "eml", "txt"])
        if msg_file is not None:
            try:
                file_bytes = msg_file.read()
                try:
                    extracted_text = file_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    extracted_text = file_bytes.decode("euc-kr", errors="ignore")
                st.success(f"메일 파일 로드 성공: `{msg_file.name}`")
            except Exception as e:
                st.error(f"메일 파일 읽기 오류: {e}")
    else:
        sample_mails = {
            "직접 입력": "",
            "사례 1 (일정 지연 경고)": "From: 개발팀장\nSent: 2026-09-08\nSubject: RE: BSW 연동 결과 보고\n\n협력사 검증 지연으로 인해 금주 예정이었던 모듈 작성이 2주 연기되었습니다.",
            "사례 2 (요구사항 변경 이슈)": "From: PL\nSent: 2026-09-07\nSubject: FW: 고객사 변경 요청건\n\nOEM 측 요구사항 추가 반영으로 아키텍처 수정을 진행하며 추가 공수가 소요될 예정입니다.",
            "사례 3 (정상 보고)": "From: QA팀\nSent: 2026-09-06\nSubject: MS3 마일스톤 검증 승인\n\n산출물 검토 결과 특이사항 없이 승인되었습니다.",
        }
        selected_sample = st.selectbox("테스트 샘플 문진 선택:", list(sample_mails.keys()))
        extracted_text = sample_mails[selected_sample]

    st.subheader("📄 메일 추출 텍스트 확인")
    final_mail_text = st.text_area(
        "추출된 메일 텍스트 내용 (수정 및 직접 작성 가능):",
        value=extracted_text,
        height=220,
        placeholder="메일 본문 또는 .msg 파일 업로드 시 텍스트가 여기에 나타납니다.",
    )

    if st.button("🔍 메일 톤&매너 및 리스크 분석 실행", type="primary"):
        if final_mail_text.strip():
            st.markdown("---")
            st.subheader("📊 진단 분석 보고서")

            headers = {"From": "미지정", "Sent": "미지정", "Subject": "미지정"}
            for line in final_mail_text.split("\n"):
                if line.startswith("From:"): headers["From"] = line.replace("From:", "").strip()
                elif line.startswith("Sent:"): headers["Sent"] = line.replace("Sent:", "").strip()
                elif line.startswith("Subject:"): headers["Subject"] = line.replace("Subject:", "").strip()

            st.write(f"**발신자:** `{headers['From']}` | **발송일자:** `{headers['Sent']}` | **제목:** `{headers['Subject']}`")

            delay_kw = ["지연", "연기", "차질", "미뤄", "늦", "홀드"]
            scope_kw = ["요구사항", "변경", "추가", "재설계", "수정", "개정"]
            risk_kw = ["문제", "오류", "불가", "리스크", "비상", "심각", "실패", "결함"]

            f_delay = [k for k in delay_kw if k in final_mail_text]
            f_scope = [k for k in scope_kw if k in final_mail_text]
            f_risk = [k for k in risk_kw if k in final_mail_text]

            r1, r2, r3 = st.columns(3)
            r1.metric("일정 리스크", "감지됨" if f_delay else "정상", delta="경고" if f_delay else "양호", delta_color="inverse" if f_delay else "normal")
            r2.metric("요구사항 흔들림", "감지됨" if f_scope else "정상", delta="주의" if f_scope else "양호", delta_color="inverse" if f_scope else "normal")
            r3.metric("이슈/오류 어조", "감지됨" if f_risk else "정상", delta="경고" if f_risk else "양호", delta_color="inverse" if f_risk else "normal")

            st.markdown("---")
            if f_delay or f_scope or f_risk:
                st.error("🚨 **프로젝트 관리자 후속 가이드**")
                if f_delay:
                    st.write(f"- **일정 지연 어휘 감지:** `{', '.join(f_delay)}` $\\rightarrow$ EVMS 대시보드의 Schedule Variance(SV) 지표 점검 권장")
                if f_scope:
                    st.write(f"- **요구사항 변경 어휘 감지:** `{', '.join(f_scope)}` $\\rightarrow$ CCB 변경관리 심의 및 BAC/공수 재조정 필요")
                if f_risk:
                    st.write(f"- **리스크 어조 감지:** `{', '.join(f_risk)}` $\\rightarrow$ 담당자 및 협력사 긴급 점검 필요")
            else:
                st.success("✅ **양호 (Normal):** 특이 리스크 어조 및 일정/품질 위험 요소가 감지되지 않았습니다.")
        else:
            st.warning("분석할 메일 텍스트가 없습니다. .msg 파일을 업로드하시거나 텍스트를 입력해 주세요.")
