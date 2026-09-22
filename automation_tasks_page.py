# automation_tasks_page.py
# -------------------------------------------------------------
# "SW 프로젝트 관리 자동화 태스크" 목록/상세 페이지 모듈
# app_v2.py(메인 앱)에서 import 하여 사이드바 메뉴 중 하나로 호출한다.
#   사용 예:
#       import automation_tasks_page
#       ...
#       elif service_mode == "🗂️ 자동화 태스크 목록":
#           automation_tasks_page.render()
# -------------------------------------------------------------

import os
import io
import html as _html
import streamlit as st
import pandas as pd
import streamlit.components.v1 as components

# 엑셀 원본: SWPJT_Web 시트, B:G열(순/영역/태스크/설명/비고/인포그래픽), 4행 제목행, 5행부터 데이터
DEFAULT_EXCEL_FILENAME = "자동화태스크.xlsx"
SHEET_NAME = "SWPJT_Web"
HEADER_ROW_INDEX = 3   # 0-based → 엑셀 4행
USECOLS = "B:G"

# 인포그래픽(*.html) 파일이 위치한 폴더
IMAGES_FOLDER = "images"
INFOGRAPHIC_HEIGHT = 1600

# 목록 페이지 상단에 고정으로 표시할 V-Cycle 태스크 개요 다이어그램
VCYCLE_DIAGRAM_FILENAME = "0_task_vcycle_diagram_v1.html"
# 다이어그램 HTML이 폭 1200px 기준으로 높이를 스스로 계산해 고정 배치되므로
# (JS로 iframe을 재조정하지 않음 — sandbox 환경에서 신뢰할 수 없어 폐기),
# 이 값은 HTML의 실제 총 높이와 정확히 일치해야 한다.
# HTML 쪽 치수(디자인 폭 1200px, 프레임 패딩 16px, 테두리 1px, 하단 여백 12px 등)를
# 바꾸면 이 값도 함께 다시 계산해서 맞춰야 한다.
VCYCLE_DIAGRAM_HEIGHT = 769


# ---------------------------------------------------------------
# 파일 경로 탐색 (app_v2.py의 get_valid_file_path와 동일한 규칙:
# data/ 폴더 우선 → 없으면 현재 폴더)
# ---------------------------------------------------------------
def _get_valid_file_path(filename: str):
    data_folder_path = os.path.join("data", filename)
    if os.path.exists(data_folder_path):
        return data_folder_path
    elif os.path.exists(filename):
        return filename
    return None


# ---------------------------------------------------------------
# 인포그래픽(*.html) 파일 경로 탐색 (images/ 폴더 우선 → 없으면 현재 폴더)
# ---------------------------------------------------------------
def _get_infographic_path(filename: str):
    images_folder_path = os.path.join(IMAGES_FOLDER, filename)
    if os.path.exists(images_folder_path):
        return images_folder_path
    elif os.path.exists(filename):
        return filename
    return None


# ---------------------------------------------------------------
# 무캐싱 실시간 로더 (엑셀이 수정되면 바로 반영되도록 매번 새로 읽음)
# ---------------------------------------------------------------
def _load_tasks_dataframe(file_source):
    try:
        if isinstance(file_source, str):
            with open(file_source, "rb") as f:
                file_bytes = f.read()
            xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        else:
            file_source.seek(0)
            xl = pd.ExcelFile(file_source, engine="openpyxl")

        if SHEET_NAME not in xl.sheet_names:
            st.error(f"'{SHEET_NAME}' 시트를 찾을 수 없습니다. (시트 목록: {xl.sheet_names})")
            return None

        df = pd.read_excel(xl, sheet_name=SHEET_NAME, header=HEADER_ROW_INDEX, usecols=USECOLS)

        # 컬럼명 정규화: 비고(LLM 이용 외) 처럼 실제 헤더 문구가 바뀌어도 위치 기준으로 표준화
        df.columns = ["순", "영역", "태스크", "설명", "비고", "인포그래픽"]

        # 완전히 빈 행 제거 (태스크명이 없는 행)
        df = df.dropna(subset=["태스크"]).reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None


# ---------------------------------------------------------------
# 설명(▶/▷/ㄴ/※/- 등 다단계 불릿 텍스트)을 "원문 그대로" 표시
# ---------------------------------------------------------------
def _format_description_html(desc) -> str:
    if not isinstance(desc, str) or not desc.strip():
        return "<em>설명 없음</em>"

    text = desc.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    rendered_lines = []
    for line in text.split("\n"):
        stripped = line.lstrip(" ")
        leading = len(line) - len(stripped)
        escaped = _html.escape(stripped)
        rendered_lines.append(("&nbsp;" * leading) + escaped)

    body = "<br>".join(rendered_lines)
    return (
        '<div style="font-family:inherit; font-size:0.95rem; line-height:1.7;">'
        f'{body}'
        '</div>'
    )


def _format_note_badges(note) -> str:
    if not isinstance(note, str) or not note.strip():
        return ""
    parts = [p.strip() for p in note.replace("\r\n", "\n").split("\n") if p.strip()]
    return "  ".join(f"`{p}`" for p in parts)


# ---------------------------------------------------------------
# 인포그래픽 셀 값 파싱: 콤마로 구분된 여러 파일명을 지원한다.
# ---------------------------------------------------------------
def _parse_infographic_list(raw_value) -> list:
    if not isinstance(raw_value, str) or not raw_value.strip():
        return []
    normalized = raw_value.replace("\r\n", "\n")
    parts = normalized.split(",")
    names = []
    for part in parts:
        name = part.replace("\n", "").strip()
        if name:
            names.append(name)
    return names


# ---------------------------------------------------------------
# 상세 페이지
# ---------------------------------------------------------------
def _render_detail(df: pd.DataFrame, task_id):
    row_match = df[df["순"] == task_id]
    if row_match.empty:
        st.warning("선택한 태스크를 찾을 수 없습니다. 목록으로 돌아갑니다.")
        st.session_state["atk_selected_task"] = None
        st.rerun()
        return

    row = row_match.iloc[0]

    if st.button("← 전체 목록으로", key="atk_back_btn"):
        st.session_state["atk_selected_task"] = None
        st.rerun()

    st.markdown("---")
    st.caption(f"#{row['순']} · {row['영역']}")
    st.title(str(row["태스크"]).lstrip("•").strip())

    badges = _format_note_badges(row["비고"])
    if badges:
        st.markdown(f"**사용 기술/비고:** {badges}")

    st.markdown("---")
    st.subheader("📋 핵심 기능")
    st.markdown(_format_description_html(row["설명"]), unsafe_allow_html=True)

    infographic_files = _parse_infographic_list(row["인포그래픽"])
    if infographic_files:
        st.markdown("---")
        st.subheader("🖼️ 인포그래픽")
        total = len(infographic_files)
        for idx, infographic_name in enumerate(infographic_files, start=1):
            infographic_path = _get_infographic_path(infographic_name)
            if infographic_path:
                if total > 1:
                    st.caption(f"{idx}/{total} · {infographic_name}")
                with open(infographic_path, "r", encoding="utf-8") as f:
                    html_content = f.read()
                components.html(html_content, height=INFOGRAPHIC_HEIGHT, scrolling=True)
            else:
                st.info(f"인포그래픽 파일을 찾을 수 없습니다: `{IMAGES_FOLDER}/{infographic_name}`")


# ---------------------------------------------------------------
# 목록 페이지
# ---------------------------------------------------------------
def _render_list(df: pd.DataFrame):
    st.title("🗂️ SW 개발을 포함하는 프로젝트 엔지니어링 및 관리 업무 자동화 태스크")
    st.caption(f"총 {len(df)}개 태스크 · 자동화태스크.xlsx / {SHEET_NAME} 시트 연동")

    vcycle_path = _get_infographic_path(VCYCLE_DIAGRAM_FILENAME)
    if vcycle_path:
        with open(vcycle_path, "r", encoding="utf-8") as f:
            vcycle_html = f.read()
        components.html(vcycle_html, height=VCYCLE_DIAGRAM_HEIGHT, scrolling=False)

        # V-Cycle 다이어그램 다운로드 기능 추가 (html, svg, png, jpg 포맷 선택 지원)
        base_name, _ = os.path.splitext(VCYCLE_DIAGRAM_FILENAME)
        available_formats = []
        for ext in ["html", "svg", "png", "jpg"]:
            candidate_name = f"{base_name}.{ext}"
            if _get_infographic_path(candidate_name):
                available_formats.append(ext)
        if not available_formats:
            available_formats = ["html"]

        col_dl1, col_dl2 = st.columns([1, 2])
        with col_dl1:
            dl_format = st.selectbox("다운로드 형식 선택", available_formats, key="vcycle_dl_format")

        target_dl_filename = f"{base_name}.{dl_format}"
        target_dl_path = _get_infographic_path(target_dl_filename)
        if target_dl_path and os.path.exists(target_dl_path):
            file_mime_map = {
                "html": "text/html",
                "svg": "image/svg+xml",
                "png": "image/png",
                "jpg": "image/jpeg"
            }
            with open(target_dl_path, "rb") as dl_f:
                file_bytes = dl_f.read()
            with col_dl2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                st.download_button(
                    label=f"📥 V-Cycle 다이어그램 다운로드 ({dl_format.upper()})",
                    data=file_bytes,
                    file_name=target_dl_filename,
                    mime=file_mime_map.get(dl_format, "application/octet-stream"),
                    key="vcycle_download_btn"
                )

    area_options = ["전체"] + sorted(df["영역"].dropna().unique().tolist())
    selected_area = st.sidebar.selectbox("🔎 업무영역 필터", area_options)

    keyword = st.text_input("🔍 태스크명 검색", placeholder="예: TARA, 해시, JIRA ...")

    filtered = df.copy()
    if selected_area != "전체":
        filtered = filtered[filtered["영역"] == selected_area]
    if keyword.strip():
        filtered = filtered[filtered["태스크"].str.contains(keyword.strip(), case=False, na=False)]

    st.markdown("---")

    if filtered.empty:
        st.info("조건에 맞는 태스크가 없습니다.")
        return

    for area in filtered["영역"].dropna().unique():
        st.subheader(f"📁 {area}")
        area_df = filtered[filtered["영역"] == area]

        for _, row in area_df.iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([5, 1])
                with c1:
                    st.markdown(f"**#{row['순']} · {str(row['태스크']).lstrip('•').strip()}**")
                    badges = _format_note_badges(row["비고"])
                    if badges:
                        st.caption(badges)
                with c2:
                    if st.button("자세히 보기", key=f"atk_detail_btn_{row['순']}"):
                        st.session_state["atk_selected_task"] = row["순"]
                        st.rerun()
        st.markdown("")


# ---------------------------------------------------------------
# 외부(app_v2.py)에서 호출하는 진입점
# ---------------------------------------------------------------
def render():
    if "atk_selected_task" not in st.session_state:
        st.session_state["atk_selected_task"] = None

    st.sidebar.header("🗂️ 자동화 태스크 데이터")
    uploaded_file = st.sidebar.file_uploader(
        "또는 보유하신 자동화태스크 엑셀 파일 업로드", type=["xlsx"], key="atk_uploader"
    )

    if uploaded_file is not None:
        target_file = uploaded_file
        display_name = uploaded_file.name
        file_valid = True
        current_mtime = getattr(uploaded_file, "size", 0)
    else:
        target_file = _get_valid_file_path(DEFAULT_EXCEL_FILENAME)
        display_name = target_file if target_file else DEFAULT_EXCEL_FILENAME
        file_valid = target_file is not None and os.path.exists(target_file)
        current_mtime = os.path.getmtime(target_file) if file_valid else 0

    if "atk_last_mtime" not in st.session_state:
        st.session_state["atk_last_mtime"] = current_mtime
    elif st.session_state["atk_last_mtime"] != current_mtime:
        st.session_state["atk_last_mtime"] = current_mtime
        st.rerun()

    if not file_valid:
        st.warning(
            f"경고: `{DEFAULT_EXCEL_FILENAME}` 파일이 `data/` 폴더 또는 메인 폴더에 존재하지 않습니다. "
            "사이드바에서 파일을 업로드해 주세요."
        )
        return

    st.success(f"현재 연동 파일: **{display_name}**")

    df = _load_tasks_dataframe(target_file)
    if df is None or df.empty:
        return

    if st.session_state["atk_selected_task"] is not None:
        _render_detail(df, st.session_state["atk_selected_task"])
    else:
        _render_list(df)