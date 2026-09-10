# kpms_schedule_page.py
# -------------------------------------------------------------
# KPMS "테일러링" 시트 기반 프로세스 테일러링 + 일정 자동생성 기능의
# 순수 Python 재구현 (데모용).
#
# 원본: KPMS_template_v1.xlsm 의 VBA 모듈 mSchedule.bas
#   - Public Sub BuildScheduleSheet()  (테일러링 → 일정 시트 골격 생성)
#   - Public Sub DrawGanttBars()       (E열에 수기 입력한 시작/종료일로 간트 표시)
#
# [설계 결정 / 원본과 다른 점]
#   1) 원본은 "테일러링" 시트에서 필터링한 Task 목록만 '일정' 시트에 만들어두고,
#      실제 시작/종료일은 사용자가 E열에 하나하나 수기 입력한 뒤 DrawGanttBars를
#      별도 실행해야 간트 막대가 그려진다. 이번 데모는 "자동 생성"이 목적이므로,
#      각 Task에 대해 소속 마일스톤 구간 안에서 균등 분배한 시작/종료일을 자동으로
#      배정해서 간트 차트까지 한 번에 보여준다 (수기 조정 단계는 생략).
#   2) 도형/셀 병합 기반 레이아웃 대신 Plotly 타임라인(Gantt) 차트로 시각화한다.
#   3) 사용자가 요청한 6가지 옵션(최초개발/파생개발/버전개발/국책과제/자체과제/기타)은
#      원본 시트의 8개 컬럼(표준/수주1~4/자체/국책/직접지정)과 정확히 1:1로 대응하지
#      않는다. 아래와 같이 매핑했다 — 특히 "버전개발"은 원본에 대응 컬럼이 없어
#      "표준(Full-Set)" 컬럼으로 대체했다(가장 보수적으로 전체 항목을 포함하는 기준).
#        최초개발 -> 수주1(TC_S1)   파생개발 -> 수주2(TC_S2)
#        버전개발 -> 표준(TC_STD, 대응 컬럼 없어 대체)
#        국책과제 -> 국책과제(TC_GOV)   자체과제 -> 자체과제(TC_SELF)
#        기타     -> 직접지정(전체 포함)
#   4) 마일스톤 구간 판정 기호는 원본과 동일하게 "●"(수행) 또는 "▲"(조건부)만 포함,
#      "×"·공란은 제외.
# -------------------------------------------------------------

import os
import io
from datetime import date

import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.utils import column_index_from_string
from dateutil.relativedelta import relativedelta

DEFAULT_EXCEL_FILENAME = "KPMS_template_v1.xlsm"
SHT_TAILORING = "테일러링"
SHT_TERM = "용어"

TAILORING_ROW_FIRST = 5

# "테일러링" 시트 컬럼 (mSchedule.bas 상수와 동일한 위치)
COL_NO = "B"
COL_MS = "C"
COL_VC = "D"
COL_TASK = "E"

# 6가지 데모 옵션 -> 테일러링 시트 컬럼 매핑 (모듈 상단 주석 참고)
TAILORING_OPTIONS = {
    "최초개발":  {"col": "K", "label": "수주1-최초개발", "desc": "SW 전체 개발"},
    "파생개발":  {"col": "L", "label": "수주2-파생개발", "desc": "SW 일부 개발"},
    "버전개발":  {"col": "J", "label": "표준(Full-Set) 대체", "desc": "원본에 대응 컬럼 없어 표준 기준 대체"},
    "국책과제":  {"col": "P", "label": "국책과제", "desc": "정부 지원 R&D"},
    "자체과제":  {"col": "O", "label": "자체과제", "desc": "자체 투자 R&D"},
    "기타":      {"col": None, "label": "직접지정", "desc": "전체 Task 포함"},
}

MILESTONE_DEFAULT_DUR = {
    "착수 ~ M-CAR": 6,
    "M-CAR ~ Proto": 12,
    "Proto ~ P1": 12,
    "P1 ~ P2": 12,
    "P2 ~ SoP": 13,
    "SoP ~": 6,
}
MILESTONE_ORDER = list(MILESTONE_DEFAULT_DUR.keys())

VC_COLOR = {
    "프로젝트 관리":    "#D9D9D9",
    "고객요구조건":     "#BDD7EE",
    "시스템 아키텍처":  "#DEEAF1",
    "SW 아키텍처":      "#D9EAD3",
    "SW 초기 개발":     "#C6E0B4",
    "컴포넌트 설계":    "#C6E0B4",
    "유닛 설계":        "#A9D18E",
    "안전/사이버보안":  "#FCE4D6",
    "유닛 검증":        "#FFF2CC",
    "상세설계 검증":    "#FFE699",
    "시스템 통합 검증": "#E2D0F0",
    "차량 검증":        "#F4CCCC",
    "릴리즈":           "#1F3864",
    "양산 후 관리":     "#F0E6D3",
}
VC_COLOR_DEFAULT = "#F2F2F2"


# =================================================================
# 공용 유틸
# =================================================================
def _get_valid_file_path(filename: str):
    data_folder_path = os.path.join("data", filename)
    if os.path.exists(data_folder_path):
        return data_folder_path
    elif os.path.exists(filename):
        return filename
    return None


def _col(letter):
    return column_index_from_string(letter)


def _cell(ws, row, letter):
    return ws.cell(row=row, column=_col(letter)).value


def _load_workbook(file_source):
    if isinstance(file_source, str):
        with open(file_source, "rb") as f:
            file_bytes = f.read()
    else:
        file_source.seek(0)
        file_bytes = file_source.read()
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, keep_vba=False)


# =================================================================
# 마일스톤 기간 (용어 시트, 없으면 기본값)
# =================================================================
def _get_milestone_durations(wb):
    durations = dict(MILESTONE_DEFAULT_DUR)
    if SHT_TERM in wb.sheetnames:
        ws = wb[SHT_TERM]
        for r in range(5, 16):
            name = _cell(ws, r, "A")
            if name is None:
                continue
            name = str(name).strip()
            if name in durations:
                hv = _cell(ws, r, "H")
                if isinstance(hv, (int, float)) and hv > 0:
                    durations[name] = int(hv)
    return durations


# =================================================================
# 테일러링 시트 -> 필터링된 Task 목록
# =================================================================
def _get_tailoring_tasks(ws, tail_col_letter):
    tasks = []
    last_row = ws.max_row
    for r in range(TAILORING_ROW_FIRST, last_row + 1):
        b_val = _cell(ws, r, COL_NO)
        if b_val is None or not isinstance(b_val, (int, float)):
            continue
        task_no = int(b_val)
        task_ms = str(_cell(ws, r, COL_MS) or "").strip()
        task_vc = str(_cell(ws, r, COL_VC) or "").strip()
        task_name = str(_cell(ws, r, COL_TASK) or "").strip()

        if "~" not in task_ms:
            continue

        if tail_col_letter is None:
            included = True
        else:
            sym = str(_cell(ws, r, tail_col_letter) or "").strip()
            included = sym in ("●", "▲")

        if not included:
            continue

        tasks.append({"no": task_no, "ms": task_ms, "vc": task_vc, "task": task_name})

    return tasks


# =================================================================
# 마일스톤 구간 내 Task 자동 날짜 배정 (균등 분배 데모)
# =================================================================
def _assign_schedule(tasks, milestone_durations, start_date):
    # 마일스톤별 시작/종료월 계산
    ms_bounds = {}
    cursor = start_date
    for ms in MILESTONE_ORDER:
        dur = milestone_durations.get(ms, MILESTONE_DEFAULT_DUR[ms])
        if dur <= 0:
            continue
        ms_start = cursor
        ms_end = cursor + relativedelta(months=dur) - relativedelta(days=1)
        ms_bounds[ms] = (ms_start, ms_end, dur)
        cursor = cursor + relativedelta(months=dur)

    # 마일스톤별로 그룹핑 (원본 순서 유지)
    by_ms = {}
    for t in tasks:
        by_ms.setdefault(t["ms"], []).append(t)

    scheduled = []
    for ms, ms_tasks in by_ms.items():
        if ms not in ms_bounds:
            continue
        ms_start, ms_end, dur = ms_bounds[ms]
        total_days = (ms_end - ms_start).days + 1
        n = len(ms_tasks)
        slice_days = max(total_days // n, 1)

        for i, t in enumerate(ms_tasks):
            seg_start = ms_start + relativedelta(days=i * slice_days)
            if i == n - 1:
                seg_end = ms_end
            else:
                seg_end = min(ms_start + relativedelta(days=(i + 1) * slice_days - 1), ms_end)
            if seg_end < seg_start:
                seg_end = seg_start
            scheduled.append({**t, "start": seg_start, "end": seg_end})

    scheduled.sort(key=lambda x: x["no"])
    return scheduled, ms_bounds


# =================================================================
# Streamlit UI
# =================================================================
def _sidebar_load_workbook():
    with st.expander("📁 KPMS 데이터 소스", expanded=False):
        uploaded_file = st.file_uploader(
            "또는 보유하신 KPMS 템플릿(.xlsm/.xlsx) 업로드", type=["xlsm", "xlsx"], key="kpms_uploader"
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

    if "kpms_last_mtime" not in st.session_state:
        st.session_state["kpms_last_mtime"] = current_mtime
    elif st.session_state["kpms_last_mtime"] != current_mtime:
        st.session_state["kpms_last_mtime"] = current_mtime
        st.rerun()

    if not file_valid:
        st.warning(
            f"경고: `{DEFAULT_EXCEL_FILENAME}` 파일이 `data/` 폴더 또는 메인 폴더에 존재하지 않습니다. "
            "사이드바에서 KPMS 템플릿 파일을 업로드해 주세요."
        )
        return None, None

    st.success(f"현재 연동 파일: **{display_name}**")
    try:
        wb = _load_workbook(target_file)
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return None, None

    if SHT_TAILORING not in wb.sheetnames:
        st.error(f"'{SHT_TAILORING}' 시트를 찾을 수 없습니다. (시트 목록: {wb.sheetnames})")
        return None, None

    return wb, display_name


def render():
    wb, _display_name = _sidebar_load_workbook()
    if wb is None:
        return

    st.title("📅 프로세스 테일러링 · 일정 자동생성 (데모)")
    st.caption("mSchedule.bas (BuildScheduleSheet / DrawGanttBars) 순수 Python 재구현")

    c1, c2 = st.columns([2, 2])
    with c1:
        option_name = st.selectbox("🏷️ 테일러링 옵션 선택", list(TAILORING_OPTIONS.keys()))
    with c2:
        start_date = st.date_input("📆 프로젝트 시작 년월", value=date.today().replace(day=1))
        start_date = start_date.replace(day=1)

    opt_cfg = TAILORING_OPTIONS[option_name]
    st.caption(f"→ 적용 기준: **{opt_cfg['label']}** ({opt_cfg['desc']})")

    if st.button("🚀 일정 자동 생성", type="primary"):
        ws_tail = wb[SHT_TAILORING]
        milestone_durations = _get_milestone_durations(wb)
        tasks = _get_tailoring_tasks(ws_tail, opt_cfg["col"])

        if not tasks:
            st.warning("선택한 옵션에 해당하는 Task를 찾지 못했습니다.")
            return

        scheduled, ms_bounds = _assign_schedule(tasks, milestone_durations, start_date)
        total_months = sum(d for _, _, d in ms_bounds.values())

        st.session_state["kpms_schedule_result"] = {
            "option_name": option_name,
            "scheduled": scheduled,
            "ms_bounds": ms_bounds,
            "total_months": total_months,
            "start_date": start_date,
        }

    result = st.session_state.get("kpms_schedule_result")
    if not result:
        return

    st.markdown("---")
    scheduled = result["scheduled"]
    ms_bounds = result["ms_bounds"]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("적용 옵션", result["option_name"])
    m2.metric("시작", result["start_date"].strftime("%Y-%m"))
    m3.metric("총 기간", f"{result['total_months']}개월")
    m4.metric("Task 수", f"{len(scheduled)}개")

    # ---- Gantt 차트 (Plotly Timeline) ----
    df = pd.DataFrame([
        {
            "번호": t["no"],
            "V-Cycle": t["vc"],
            "업무": f"[{t['no']}] {t['task']}",
            "마일스톤": t["ms"],
            "시작": pd.Timestamp(t["start"]),
            "종료": pd.Timestamp(t["end"] + relativedelta(days=1)),  # Plotly는 종료일 배타적이라 +1일 보정
        }
        for t in scheduled
    ])

    color_map = {vc: VC_COLOR.get(vc, VC_COLOR_DEFAULT) for vc in df["V-Cycle"].unique()}

    fig = px.timeline(
        df.sort_values("번호", ascending=False),
        x_start="시작", x_end="종료", y="업무", color="V-Cycle",
        color_discrete_map=color_map,
        hover_data={"마일스톤": True, "번호": True, "업무": False},
    )
    fig.update_yaxes(title="", tickfont=dict(size=9))
    fig.update_layout(height=max(400, 22 * len(df)), legend_title="V-Cycle",
                       margin=dict(l=10, r=10, t=30, b=10))

    # [방어 처리] 일부 plotly 버전은 날짜축 위 막대 길이를 datetime.timedelta로
    # 내부 표현하는데, 그 값을 JSON으로 직렬화하는 로직이 없어 오류가 나는 경우가
    # 있다(구버전 plotly의 알려진 버그). 막대 길이를 밀리초 단위 float로 직접
    # 변환해 어떤 plotly 버전에서도 안전하게 직렬화되도록 만든다.
    for trace in fig.data:
        x_vals = getattr(trace, "x", None)
        if x_vals is None:
            continue
        sanitized = []
        changed = False
        for v in x_vals:
            total_seconds = None
            if isinstance(v, pd.Timedelta):
                total_seconds = v.total_seconds()
            elif hasattr(v, "total_seconds"):
                total_seconds = v.total_seconds()
            if total_seconds is not None:
                sanitized.append(total_seconds * 1000.0)
                changed = True
            else:
                sanitized.append(v)
        if changed:
            trace.x = sanitized

    # 마일스톤 구간 경계 표시
    for ms, (ms_start, ms_end, _dur) in ms_bounds.items():
        fig.add_vline(x=pd.Timestamp(ms_start), line_width=1, line_dash="dot", line_color="gray")

    st.subheader("📊 자동 생성된 프로젝트 일정 (Gantt)")
    st.plotly_chart(fig, use_container_width=True)

    # ---- 마일스톤 구간 요약 ----
    st.subheader("🚩 마일스톤 구간")
    ms_df = pd.DataFrame([
        {"마일스톤": ms, "시작": s.strftime("%Y-%m-%d"), "종료": e.strftime("%Y-%m-%d"), "기간(개월)": d}
        for ms, (s, e, d) in ms_bounds.items()
    ])
    st.dataframe(ms_df, use_container_width=True)

    # ---- Task 목록 테이블 ----
    st.subheader("📋 Task 목록")
    task_df = pd.DataFrame([
        {
            "번호": t["no"], "마일스톤": t["ms"], "V-Cycle": t["vc"], "업무(Task)": t["task"],
            "시작일": t["start"].strftime("%Y-%m-%d"), "종료일": t["end"].strftime("%Y-%m-%d"),
        }
        for t in scheduled
    ])
    st.dataframe(task_df, use_container_width=True)

    st.download_button(
        "📥 Task 목록 CSV 다운로드",
        data=task_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"KPMS_schedule_{result['option_name']}.csv",
        mime="text/csv",
    )
