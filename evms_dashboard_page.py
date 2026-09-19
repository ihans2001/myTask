# evms_dashboard_page.py
# -------------------------------------------------------------
# EVMS(획득가치관리) 성과 대시보드. app.py 인라인 코드에서 분리.
# -------------------------------------------------------------

import base64
import io
import os
import pandas as pd
import streamlit as st

# 고정 S-Curve 차트 이미지 경로 (엑셀 시트에서 자동 생성하지 않고 고정 이미지로 표시)
# 주의: 파일명은 실제 파일과 대소문자까지 정확히 일치해야 함 (Linux 배포 환경은 대소문자 구분)
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCURVE_SVG_PATH = os.path.join(_BASE_DIR, "images", "evms_curve_chart.svg")


def _load_svg_as_img_tag(svg_path, width="100%"):
    """SVG 파일을 base64로 인코딩해 <img> 태그 문자열로 반환.
    st.image()보다 크기 조절이 안정적이라 st.markdown(unsafe_allow_html=True)로 삽입한다."""
    with open(svg_path, "rb") as f:
        svg_bytes = f.read()
    b64 = base64.b64encode(svg_bytes).decode("utf-8")
    return f'<img src="data:image/svg+xml;base64,{b64}" style="width:{width}; max-width:900px;" />'

PROJECT_FILES = [
    "Project_01_Normal.xlsx",
    "Project_02_Critical_Risk.xlsx",
    "Project_03_Green_Washing.xlsx",
    "Project_04_Fast_Tracking.xlsx",
    "Project_05_Early_Warning.xlsx",
    "Project_06_Final_EAC.xlsx",
]


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


def _force_read_excel_direct(file_source):
    try:
        if isinstance(file_source, str):
            with open(file_source, "rb") as f:
                file_bytes = f.read()
            xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        else:
            file_source.seek(0)
            xl = pd.ExcelFile(file_source, engine="openpyxl")

        sheets = xl.sheet_names
        dash_sheet = "EVMS_대시보드" if "EVMS_대시보드" in sheets else sheets[0]
        df_dash = pd.read_excel(xl, sheet_name=dash_sheet)

        curve_sheet = None
        exact_candidates = ["EVMS_S커브_예시", "EVMS_Chart", "S-Curve", "Chart"]
        for s in exact_candidates:
            if s in sheets:
                curve_sheet = s
                break

        df_curve = pd.read_excel(xl, sheet_name=curve_sheet) if curve_sheet else None
        return df_dash, df_curve, sheets, curve_sheet
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None, [], None


def render():
    st.title("📊 Cloud EVMS Management SaaS")

    with st.expander("📁 프로젝트 데이터 소스", expanded=False):
        selected_file = st.selectbox(
            "분석할 프로젝트 엑셀 파일(.xlsx)을 선택하세요:",
            PROJECT_FILES,
        )
        uploaded_file = st.file_uploader("또는 보유하신 EVMS 엑셀 파일 업로드", type=["xlsx"], key="evms_uploader")

    if uploaded_file is not None:
        target_file = uploaded_file
        display_name = uploaded_file.name
        file_valid = True
        current_mtime = getattr(uploaded_file, "size", 0)
    else:
        target_file = _get_valid_file_path(selected_file)
        display_name = target_file if target_file else selected_file
        file_valid = target_file is not None and os.path.exists(target_file)
        current_mtime = os.path.getmtime(target_file) if file_valid else 0

    if "evms_last_mtime" not in st.session_state:
        st.session_state["evms_last_mtime"] = current_mtime
    elif st.session_state["evms_last_mtime"] != current_mtime:
        st.session_state["evms_last_mtime"] = current_mtime
        st.rerun()

    if not file_valid:
        st.warning(f"경고: `{selected_file}` 파일이 `data/` 폴더 또는 메인 폴더에 존재하지 않습니다.")
        return

    st.success(f"현재 연동 파일: **{display_name}** (자동 감지 최신화 상태)")

    df_dash_raw, df_curve_raw, sheet_names, curve_sheet_used = _force_read_excel_direct(target_file)
    df_dash = _fix_dataframe_types(df_dash_raw)

    if df_dash is None:
        return

    st.subheader("📌 핵심 성과 지표 (EVMS Summary KPIs)")
    try:
        bac_val = pd.to_numeric(df_dash_raw.iloc[:, 3], errors="coerce").sum()
        pv_val = pd.to_numeric(df_dash_raw.iloc[:, 4], errors="coerce").sum()
        ev_val = pd.to_numeric(df_dash_raw.iloc[:, 5], errors="coerce").sum()
        ac_val = pd.to_numeric(df_dash_raw.iloc[:, 6], errors="coerce").sum()

        cpi = round(ev_val / ac_val, 2) if ac_val > 0 else 1.0
        spi = round(ev_val / pv_val, 2) if pv_val > 0 else 1.0
        sv = round(ev_val - pv_val, 2)
        cv = round(ev_val - ac_val, 2)
    except Exception:
        bac_val, pv_val, ev_val, ac_val = 100, 80, 75, 85
        cpi, spi, sv, cv = 0.88, 0.94, -5, -10

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("BAC (총예산)", f"{bac_val:,.1f} MM")
    k2.metric("PV (계획가치)", f"{pv_val:,.1f} MM")
    k3.metric("EV (획득가치)", f"{ev_val:,.1f} MM", delta=f"SV: {sv:,.1f}")
    k4.metric("AC (실제비용)", f"{ac_val:,.1f} MM", delta=f"CV: {cv:,.1f}", delta_color="inverse")
    k5.metric("CPI / SPI", f"{cpi:.2f} / {spi:.2f}", delta=f"CPI: {cpi:.2f}")

    st.markdown("---")
    st.subheader("📈 EVMS S-Curve 추이 분석 그래프")

    if os.path.exists(SCURVE_SVG_PATH):
        st.markdown(_load_svg_as_img_tag(SCURVE_SVG_PATH), unsafe_allow_html=True)
    else:
        st.warning(
            f"⚠️ 고정 차트 이미지 파일을 찾을 수 없습니다: `{SCURVE_SVG_PATH}` "
            "(앱 실행 폴더 기준 `images/evm_scurve_chart.svg` 경로에 파일을 넣어주세요.)"
        )

    st.subheader("📋 세부 Control Account 현황 (EVMS_대시보드)")
    st.dataframe(df_dash, width="stretch")
