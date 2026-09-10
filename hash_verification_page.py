# hash_verification_page.py
# -------------------------------------------------------------
# SW 산출물 해시 검증 (verifyHash). app.py 인라인 코드에서 분리.
# -------------------------------------------------------------

import os
import subprocess
import hashlib
import pandas as pd
import streamlit as st


def _get_valid_file_path(filename):
    data_folder_path = os.path.join("data", filename)
    if os.path.exists(data_folder_path):
        return data_folder_path
    elif os.path.exists(filename):
        return filename
    return None


def _fix_dataframe_types(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    df.columns = [str(c) if c is not None else "" for c in df.columns]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).fillna("")
    return df


def render():
    st.title("🔐 SW 산출물 위변조 및 해시 검증 모듈")
    st.caption("verify_hash_each.py 및 verifyHash.xlsm 브릿지 연동")

    st.info("💡 검증 단위(파일 단위 / 엑셀 워크시트 단위)와 알고리즘을 선택한 후 해시 계산 및 Diff 대조를 수행합니다.")

    tab1, tab2 = st.tabs(["📄 파일 / 워크시트 단위 해시 검증 & Diff 대조", "⚙️ 로컬 파이썬 브릿지 스크립트 실행"])

    with tab1:
        c_unit, c_algo, c_chunk = st.columns([2, 2, 1.5])
        with c_unit:
            calc_unit = st.radio(
                "🎯 해시 계산 단위 선택:",
                ["📁 파일 단위 해시 계산", "📊 엑셀 워크시트(Worksheet) 단위 해시 계산"],
                horizontal=False,
            )
        with c_algo:
            hash_algo = st.selectbox(
                "🔑 암호화/해시 알고리즘 선택:",
                ["SHA-256 (표준)", "SHA-1", "MD5", "SHA-512"],
            )
        with c_chunk:
            chunk_mode = st.checkbox("64KB Chunk 계산 적용", value=True)

        st.markdown("---")
        st.subheader("📁 해시 대상 파일 업로드 (Drag & Drop 또는 Browse)")

        allowed_types = ["xlsx", "xlsm", "xls"] if "워크시트" in calc_unit else None
        uploaded_hash_files = st.file_uploader(
            "검증 대상 SW 산출물 파일들을 업로드하세요:" if "파일" in calc_unit else "워크시트별 해시를 산출할 엑셀 파일(.xlsx, .xlsm)을 업로드하세요:",
            accept_multiple_files=True,
            type=allowed_types,
        )

        st.subheader("🔍 해시 Diff 대조 (선택 사항)")
        expected_hash_input = st.text_input(
            "대조할 원본/기대 해시값(Expected Hash)을 입력하세요 (여러 항목인 경우 쉼표 또는 줄바꿈으로 구분):",
            placeholder="예: 4A7B... 또는 64자리 16진수 문자열",
        )

        if uploaded_hash_files:
            st.subheader("📊 해시 계산 및 Diff 대조 결과")
            expected_list = [h.strip().upper() for h in expected_hash_input.replace("\n", ",").split(",") if h.strip()]

            results = []

            if "파일 단위" in calc_unit:
                for idx, file_item in enumerate(uploaded_hash_files):
                    if "SHA-1" in hash_algo: hasher = hashlib.sha1()
                    elif "MD5" in hash_algo: hasher = hashlib.md5()
                    elif "SHA-512" in hash_algo: hasher = hashlib.sha512()
                    else: hasher = hashlib.sha256()

                    file_bytes = file_item.read()
                    if chunk_mode:
                        chunk_size = 65536
                        for i in range(0, len(file_bytes), chunk_size):
                            hasher.update(file_bytes[i:i + chunk_size])
                    else:
                        hasher.update(file_bytes)

                    calc_hash = hasher.hexdigest().upper()
                    exp_val = expected_list[idx] if idx < len(expected_list) else "-"

                    match_status = "✅ MATCH (일치)" if exp_val != "-" and calc_hash == exp_val else ("🚨 MISMATCH (위변조 감지)" if exp_val != "-" else "➖ (기대값 미입력)")

                    results.append({
                        "대상 명칭 (파일명 / 시트명)": file_item.name,
                        "검증 단위": "파일 (File)",
                        "크기 (Byte)": len(file_bytes),
                        "알고리즘": hash_algo.split(" ")[0],
                        "계산된 해시값 (Calculated)": calc_hash,
                        "기대 해시값 (Expected)": exp_val,
                        "Diff 검증 결과": match_status,
                    })
            else:
                item_count = 0
                for file_item in uploaded_hash_files:
                    try:
                        xl = pd.ExcelFile(file_item, engine="openpyxl")
                        for sheet_name in xl.sheet_names:
                            df_sheet = pd.read_excel(xl, sheet_name=sheet_name)
                            sheet_bytes = df_sheet.to_csv(index=False).encode("utf-8")

                            if "SHA-1" in hash_algo: hasher = hashlib.sha1()
                            elif "MD5" in hash_algo: hasher = hashlib.md5()
                            elif "SHA-512" in hash_algo: hasher = hashlib.sha512()
                            else: hasher = hashlib.sha256()

                            hasher.update(sheet_bytes)
                            calc_hash = hasher.hexdigest().upper()

                            exp_val = expected_list[item_count] if item_count < len(expected_list) else "-"
                            match_status = "✅ MATCH (일치)" if exp_val != "-" and calc_hash == exp_val else ("🚨 MISMATCH (위변조 감지)" if exp_val != "-" else "➖ (기대값 미입력)")

                            results.append({
                                "대상 명칭 (파일명 / 시트명)": f"{file_item.name} -> [{sheet_name}]",
                                "검증 단위": "워크시트 (Worksheet)",
                                "크기 (Byte)": len(sheet_bytes),
                                "알고리즘": hash_algo.split(" ")[0],
                                "계산된 해시값 (Calculated)": calc_hash,
                                "기대 해시값 (Expected)": exp_val,
                                "Diff 검증 결과": match_status,
                            })
                            item_count += 1
                    except Exception as e:
                        st.error(f"`{file_item.name}` 워크시트 해시 계산 에러: {e}")

            df_res = _fix_dataframe_types(pd.DataFrame(results))
            st.dataframe(df_res, use_container_width=True)

    with tab2:
        st.subheader("로컬 `verify_hash_each.py` 파이썬 브릿지 실행")
        script_path = _get_valid_file_path("verify_hash_each.py")
        if script_path:
            st.success(f"연동 파이썬 스크립트 위치 확인됨: `{script_path}`")
            col_a, col_b = st.columns(2)
            with col_a:
                cmd_option = st.radio("스크립트 실행 파라미터 선택:", ["--mode file (파일 단위)", "--mode sheet (워크시트 단위)"])
            with col_b:
                if st.button("🚀 파이썬 해시 검증 스크립트 실행"):
                    mode_flag = "sheet" if "sheet" in cmd_option else "file"
                    try:
                        res = subprocess.run(["python", script_path, "--mode", mode_flag], capture_output=True, text=True, check=True)
                        st.code(res.stdout if res.stdout else "검증 완료: 출력 내용 정상", language="text")
                    except Exception as e:
                        st.error(f"스크립트 실행 에러: {e}")
        else:
            st.warning("`verify_hash_each.py` 파일이 메인 폴더 또는 data/ 폴더에 존재하지 않습니다.")
