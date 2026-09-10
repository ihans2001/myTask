# dfa_fta_page.py
# -------------------------------------------------------------
# DFA(종속고장분석)/FTA(고장수목분석) 다이어그램 자동 생성 모듈
# 원본: diagramFault.xlsm 의 VBA 모듈 mDiagramFault.bas (GenerateFaultDiagram)
#
# 원본 주석 그대로: "고장원인이 1개면 DFA, 여러 개면 FTA로 동일 로직 사용"
# -> 하나의 알고리즘으로 두 다이어그램을 모두 생성한다:
#    ① 계층 구조 트리 (제어기/SW 아키텍처를 Lvl1~Lvl9 아웃라인 표로 표현)
#    ② 고장원인 마름모 (여러 리프가 같은 (inElm, inVar) 조합에 의존하면,
#       그 공통 의존성을 마름모로 묶어 표시 -> 종속고장/공통원인고장 시각화)
#
# [Python 포팅 시 설계 결정]
#   1) Excel Shape 좌표 계산 대신 Graphviz 자동 레이아웃을 사용한다. 노드 판정
#      (Root/중간/Leaf), 레벨별 그라데이션 색상, 마름모 그룹핑 로직은 원본과
#      동일하게 재현했다.
#   2) 원본 VBA는 funcSSoT 시트의 C열/E열을 고정 인덱스로 읽어 기능 설명을
#      매칭하는데, 이는 특정 예시 파일의 레이아웃에만 맞는 방식으로 보인다.
#      Python 버전은 같은 행 번호(row)를 기준으로 diagramFTA/DFA와 funcSSoT의
#      "그 행의 최심 노드 설명"을 매칭하는 방식으로 일반화했다(두 시트가 같은
#      계층을 행 단위로 나란히 표현한다는 전제).
#   3) VBA의 HSL→RGB 변환은 Python 표준 라이브러리 colorsys로 대체했다(동일한
#      결과를 내는 표준 알고리즘).
# -------------------------------------------------------------

import os
import io
import colorsys

import streamlit as st
import openpyxl
import graphviz

DEFAULT_EXCEL_FILENAME = "diagramFault.xlsm"
MAX_LEVEL_COLS = 9          # Lvl1~Lvl9
HIERARCHY_START_ROW = 5
FUNC_SHEET_NAME = "funcSSoT"

# 후보 데이터 시트(원본 3개 탭 - 구조는 동일, 관점만 다름)
CANDIDATE_SHEETS = ["diagramFTA", "diagramDFA", "diagramFault"]

# 레벨 그라데이션 색상 (Lvl1=진한 네이비 -> 최심레벨=밝은 하늘색), mDiagramFault.bas GetLevelColor 동일
LEVEL_COLOR_START = (31, 58, 82)
LEVEL_COLOR_END = (219, 229, 236)

# 마름모(고장원인 그룹) 색상 - GetGroupColor 동일(주황 계열 Hue=32도 고정)
GROUP_HUE_DEG = 32


# ---------------------------------------------------------------
# 공용 유틸
# ---------------------------------------------------------------
def _get_valid_file_path(filename: str):
    data_folder_path = os.path.join("data", filename)
    if os.path.exists(data_folder_path):
        return data_folder_path
    elif os.path.exists(filename):
        return filename
    return None


def _load_workbook(file_source):
    if isinstance(file_source, str):
        with open(file_source, "rb") as f:
            file_bytes = f.read()
    else:
        file_source.seek(0)
        file_bytes = file_source.read()
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, keep_vba=False)


def _rgb_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(v)))) for v in rgb)


def _get_level_color(level, max_level):
    """레벨(1~max_level)에 따라 네이비->하늘색 그라데이션 색상을 hex로 반환."""
    total_steps = max_level - 1
    ratio = 0.0 if total_steps <= 0 else (level - 1) / total_steps
    ratio = min(1.0, max(0.0, ratio))
    rgb = [s + (e - s) * ratio for s, e in zip(LEVEL_COLOR_START, LEVEL_COLOR_END)]
    return _rgb_to_hex(rgb)


def _get_contrast_text_color(bg_hex):
    r = int(bg_hex[1:3], 16)
    g = int(bg_hex[3:5], 16)
    b = int(bg_hex[5:7], 16)
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    return "#ffffff" if luminance < 140 else "#333333"


def _get_group_color(idx, total):
    """마름모(고장원인 그룹) 색상 - 주황 계열, 그룹마다 채도/명도를 살짝 변조."""
    hue = GROUP_HUE_DEG / 360.0
    if total <= 1:
        sat, lum = 0.55, 0.82
    else:
        ratio = idx / (total - 1)
        sat = 0.45 + 0.25 * ratio
        lum = 0.85 - 0.25 * ratio
    r, g, b = colorsys.hls_to_rgb(hue, lum, sat)
    return _rgb_to_hex((r * 255, g * 255, b * 255))


# ---------------------------------------------------------------
# 계층 시트 파싱 (Lvl1~Lvl9 아웃라인 표, carry-down 방식)
# ---------------------------------------------------------------
def parse_hierarchy_sheet(ws):
    """
    diagramFTA/diagramDFA/diagramFault 시트를 파싱해 트리 구조와 리프 속성을 만든다.
    반환: (node_level, parent_of, children_of, root_order, leaf_attrs, row_deepest_name)
      - node_level: {노드명: 레벨(1~9)}
      - parent_of: {자식명: 부모명}
      - children_of: {부모명: [자식명, ...]}  (등장 순서 보존)
      - root_order: [Lvl1 노드명, ...]  (등장 순서 보존)
      - leaf_attrs: {리프명: {"outVar":..., "inVar":..., "inElm":...}}
      - row_deepest_name: {엑셀 행번호: 그 행의 최심 노드명}  (funcSSoT 매칭용)
    """
    node_level = {}
    parent_of = {}
    children_of = {}
    root_order = []
    leaf_attrs = {}
    row_deepest_name = {}

    carry = [""] * (MAX_LEVEL_COLS + 1)  # 1-indexed
    last_row = ws.max_row

    for r in range(HIERARCHY_START_ROW, last_row + 1):
        row_vals = [""] * (MAX_LEVEL_COLS + 1)
        deepest = 0
        for lv in range(1, MAX_LEVEL_COLS + 1):
            cell_val = ws.cell(row=r, column=2 + lv).value  # C=3(Lvl1) ... K=11(Lvl9)
            cell_val = str(cell_val).strip() if cell_val is not None else ""
            if cell_val:
                carry[lv] = cell_val
                deepest = lv
            row_vals[lv] = carry[lv]

        if deepest == 0:
            continue

        for lv in range(1, deepest + 1):
            nm = row_vals[lv]
            if not nm:
                continue
            if nm not in node_level:
                node_level[nm] = lv
                if lv == 1 and nm not in root_order:
                    root_order.append(nm)
            if lv > 1:
                parent_nm = row_vals[lv - 1]
                if parent_nm:
                    parent_of.setdefault(nm, parent_nm)
                    children_of.setdefault(parent_nm, [])
                    if nm not in children_of[parent_nm]:
                        children_of[parent_nm].append(nm)

        leaf_name = row_vals[deepest]
        row_deepest_name[r] = leaf_name

        out_var = str(ws.cell(row=r, column=12).value or "").strip()   # L
        in_var = str(ws.cell(row=r, column=13).value or "").strip()    # M
        in_elm = str(ws.cell(row=r, column=14).value or "").strip()    # N

        if leaf_name and out_var:
            leaf_attrs.setdefault(leaf_name, {"outVar": out_var, "inVar": "", "inElm": ""})
            if in_elm or in_var:
                leaf_attrs[leaf_name]["inVar"] = in_var
                leaf_attrs[leaf_name]["inElm"] = in_elm

    return node_level, parent_of, children_of, root_order, leaf_attrs, row_deepest_name


def parse_func_descriptions(ws_func, row_deepest_name_data):
    """
    funcSSoT 시트를 데이터 시트와 같은 규칙(Lvl1~Lvl9 carry-down)으로 파싱한 뒤,
    "같은 행 번호"를 기준으로 데이터 시트의 최심 노드명 -> funcSSoT의 최심 설명을 매칭한다.
    """
    if ws_func is None:
        return {}

    _, _, _, _, _, row_deepest_desc = parse_hierarchy_sheet(ws_func)
    descriptions = {}
    for r, node_name in row_deepest_name_data.items():
        desc = row_deepest_desc.get(r, "")
        if node_name and desc and node_name not in descriptions:
            descriptions[node_name] = desc
    return descriptions


# ---------------------------------------------------------------
# Graphviz 다이어그램 생성
# ---------------------------------------------------------------
def build_fault_tree_graph(node_level, parent_of, children_of, leaf_attrs, descriptions,
                            include_dfa_diamonds=True):
    max_level = max(node_level.values()) if node_level else 1

    dot = graphviz.Digraph("fault_tree")
    dot.attr(rankdir="BT")
    dot.attr("node", fontname="Noto Sans CJK KR", fontsize="10")
    dot.attr("edge", color="#787878", arrowhead="normal")

    for name, lvl in node_level.items():
        is_leaf = name in leaf_attrs
        bg = _get_level_color(lvl, max_level)
        fg = _get_contrast_text_color(bg)

        if lvl == 1:
            shape, style = "box", "rounded,filled"
            label = name
        elif is_leaf:
            shape, style = "ellipse", "filled"
            label = f"{name}\n{leaf_attrs[name]['outVar']}"
        else:
            shape, style = "box", "filled"
            desc = descriptions.get(name, "")
            label = f"{name}\n{desc}" if desc else name

        dot.node(name, label=label, shape=shape, style=style,
                  fillcolor=bg, color="#787878", fontcolor=fg)

    for child, parent in parent_of.items():
        if child in node_level and parent in node_level:
            dot.edge(child, parent)

    group_count = 0
    if include_dfa_diamonds:
        # (inElm, inVar) 쌍이 같은 리프들을 하나의 그룹(마름모)으로 묶는다
        group_members = {}
        for leaf_name, attrs in leaf_attrs.items():
            pair_key = (attrs.get("inElm", ""), attrs.get("inVar", ""))
            if pair_key == ("", ""):
                continue
            group_members.setdefault(pair_key, []).append(leaf_name)

        group_count = len(group_members)
        for gi, (pair_key, members) in enumerate(group_members.items()):
            in_elm, in_var = pair_key
            group_id = f"DIA_{gi}"
            gcolor = _get_group_color(gi, group_count)
            dot.node(group_id, label=f"{in_elm}\n{in_var}", shape="diamond",
                      style="filled", fillcolor=gcolor, color="#965A14", fontcolor="#502800")
            for leaf_name in members:
                if leaf_name in node_level:
                    dot.edge(group_id, leaf_name, color="#965A14")

    return dot, len(node_level), len(leaf_attrs), group_count


# ---------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------
def render():
    with st.expander("📁 DFA/FTA 데이터 소스", expanded=False):
        uploaded_file = st.file_uploader(
            "또는 보유하신 diagramFault 템플릿(.xlsm/.xlsx) 업로드", type=["xlsm", "xlsx"], key="dfafta_uploader"
        )

    if uploaded_file is not None:
        target_file = uploaded_file
        display_name = uploaded_file.name
        file_valid = True
    else:
        target_file = _get_valid_file_path(DEFAULT_EXCEL_FILENAME)
        display_name = target_file if target_file else DEFAULT_EXCEL_FILENAME
        file_valid = target_file is not None and os.path.exists(target_file)

    st.title("🌳 DFA/FTA 고장 다이어그램 자동 생성")
    st.caption("mDiagramFault.bas (GenerateFaultDiagram) 순수 Python 재구현 — 맥/윈도우 등 서버 환경 무관 (VBA/COM 불필요)")

    if not file_valid:
        st.warning(
            f"경고: `{DEFAULT_EXCEL_FILENAME}` 파일이 `data/` 폴더 또는 메인 폴더에 존재하지 않습니다. "
            "위 데이터 소스에서 파일을 업로드해 주세요."
        )
        return

    st.success(f"현재 연동 파일: **{display_name}**")

    try:
        wb = _load_workbook(target_file)
    except Exception as e:
        st.error(f"엑셀 파일 로드 중 오류 발생: {e}")
        return

    available_sheets = [s for s in CANDIDATE_SHEETS if s in wb.sheetnames]
    if not available_sheets:
        st.error(f"diagramFTA/diagramDFA/diagramFault 시트를 찾을 수 없습니다. (시트 목록: {wb.sheetnames})")
        return

    c1, c2 = st.columns([2, 2])
    with c1:
        sheet_name = st.selectbox("📄 데이터 시트 선택", available_sheets)
    with c2:
        include_diamonds = st.checkbox("🔶 고장원인(DFA) 마름모 함께 표시", value=True)

    if st.button("🚀 다이어그램 생성", type="primary"):
        ws = wb[sheet_name]
        node_level, parent_of, children_of, root_order, leaf_attrs, row_deepest_name = parse_hierarchy_sheet(ws)

        ws_func = wb[FUNC_SHEET_NAME] if FUNC_SHEET_NAME in wb.sheetnames else None
        descriptions = parse_func_descriptions(ws_func, row_deepest_name)

        dot, n_nodes, n_leaves, n_groups = build_fault_tree_graph(
            node_level, parent_of, children_of, leaf_attrs, descriptions,
            include_dfa_diamonds=include_diamonds,
        )
        st.session_state["dfafta_result"] = {
            "sheet_name": sheet_name,
            "dot": dot,
            "n_nodes": n_nodes,
            "n_leaves": n_leaves,
            "n_groups": n_groups,
        }

    result = st.session_state.get("dfafta_result")
    if not result:
        return

    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    m1.metric("노드 수", result["n_nodes"])
    m2.metric("리프 수", result["n_leaves"])
    m3.metric("고장원인 그룹(마름모) 수", result["n_groups"])

    st.caption(
        "🟦 색상 그라데이션(네이비→하늘색) = 레벨(Lvl1이 가장 진함) · "
        "🟧 주황 마름모 = 여러 리프가 공유하는 고장원인(공통원인고장/종속고장) · "
        "타원 = 리프(단위기능), 사각형 = 중간/루트 노드"
    )
    st.subheader(f"📊 {result['sheet_name']} 다이어그램")
    st.graphviz_chart(result["dot"], use_container_width=True)
