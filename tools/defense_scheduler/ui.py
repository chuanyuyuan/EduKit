"""答辩顺序生成器 — Streamlit UI"""
from datetime import datetime
import streamlit as st
from streamlit_sortables import sort_items
from .core import parse_names, parse_roster, build_schedule, format_duration, to_excel

SAMPLE_NAMES = """张三
李四
王五
赵六
钱七
孙八
周九
吴十
郑十一
冯十二"""

DND_CSS = """
.sortable-component { padding: 0; }
.sortable-component.vertical { gap: 0; padding: 0; }
.sortable-container { padding: 0; margin: 0; min-width: unset; flex-grow: 1; }
.sortable-container-body { padding: 0; background: transparent; border-radius: 0; min-height: auto; }
.sortable-item {
    background: #f0f2f6 !important; color: #111 !important;
    margin: 0 0 4px 0 !important; padding: 10px 14px !important;
    border-radius: 8px !important; border: 1px solid #e0e0e0;
    border-left: 4px solid #4A90D9 !important;
    cursor: grab !important; user-select: none;
    font-size: 15px; transition: background .12s;
    box-sizing: border-box;
}
.sortable-item:hover { background: #e4e7ed !important; cursor: grab !important; }
.sortable-item:active { cursor: grabbing !important; background: #d0d4dd !important; }
.sortable-item * { cursor: grab !important; }
.sortable-item:active * { cursor: grabbing !important; }
.sortable-ghost { opacity: .35; background: #A8C8F0 !important; }
"""


def _display_label(student):
    """从 (学号, 姓名) 生成拖拽组件显示文本。"""
    sid, name = student
    return f"{sid} {name}" if sid else name


def _resolve_order(sorted_displays: list[str], original: list) -> list:
    """将拖拽后的显示文本列表映射回 (学号, 姓名) 元组。"""
    lookup = {_display_label(s): s for s in original}
    return [lookup[d] for d in sorted_displays if d in lookup]


def render_page():
    st.header(":material/shuffle: 答辩顺序生成器")
    st.markdown("""
随机打乱学生答辩顺序，按总时间区间自动分配每人答辩时间，支持拖拽调整顺序。

**数据来源：**
- **上传花名册（推荐）** — 上传含学号和姓名的 Excel 花名册，自动识别「学号」「姓名」列
- **手动输入** — 在文本框中粘贴姓名列表，用逗号、顿号、空格、换行分隔均可

**操作步骤：**
1. 上传花名册或输入学生姓名
2. 设置答辩时间区间（HH:MM–HH:MM）
3. 点击「随机生成」→ 拖拽微调顺序 → 下载 Excel 答辩顺序表
""")

    # ── 文件上传（互斥于文本输入） ──
    uploaded = st.file_uploader("上传花名册（Excel）", type="xlsx")

    if uploaded:
        try:
            roster = parse_roster(uploaded)
        except ValueError as e:
            st.error(str(e))
            st.session_state.pop("order", None)
            return
        if not roster:
            st.error("未从文件中解析到学生数据，请检查是否包含「姓名」列。")
            st.session_state.pop("order", None)
            return
        st.success(f"已从花名册读取 {len(roster)} 名学生")
        with st.expander("查看学生列表", expanded=False):
            for sid, name in roster:
                st.text(f"{sid}  {name}" if sid else name)
        st.session_state.roster_data = roster
        st.session_state._use_roster = True
    else:
        st.session_state._use_roster = False

    # ── 姓名输入（无文件上传时可用） ──
    if not uploaded:
        st.text_area(
            "学生姓名",
            value=SAMPLE_NAMES,
            height=180,
            key="names_input",
            placeholder="每行一个姓名，或用逗号、顿号分隔",
        )

    col1, col2 = st.columns(2)
    with col1:
        st.text_input("开始时间", value="08:00", key="start_time")
    with col2:
        st.text_input("结束时间", value="09:30", key="end_time")

    if st.button("🎲 随机生成", type="primary", use_container_width=True):
        if uploaded:
            roster = st.session_state.get("roster_data", [])
            if not roster:
                st.error("没有花名册数据，请重新上传。")
                return
        else:
            raw = parse_names(st.session_state.names_input)
            if not raw:
                st.error("请至少输入一个姓名。")
                return
            roster = raw
        import random
        random.shuffle(roster)
        st.session_state.order = roster
        st.session_state._sort_gen = st.session_state.get("_sort_gen", 0) + 1
        st.rerun()

    if st.session_state.get("order") is not None:
        order = st.session_state.order

        st.markdown("👆 **拖拽调整顺序**")

        # 准备拖拽组件的显示文本
        display_items = [_display_label(s) for s in order]

        sorted_displays = sort_items(
            display_items,
            direction="vertical",
            custom_style=DND_CSS,
            key=f"sort_{st.session_state.get('_sort_gen', 0)}",
        )

        if sorted_displays and sorted_displays != display_items:
            st.session_state.order = _resolve_order(sorted_displays, order)
            st.rerun()

        # 重建时间表
        order = st.session_state.order
        try:
            rows, per_student = build_schedule(
                order, st.session_state.start_time, st.session_state.end_time
            )
        except ValueError as e:
            st.error(str(e))
            return

        # ── 统计信息 ──
        has_id = any(sid for sid, _ in order)
        total_min = int(per_student * len(order) // 60)
        st.success(
            f"共 {len(order)} 名学生，{st.session_state.start_time}"
            f"–{st.session_state.end_time}（{total_min} 分钟），"
            f"每人 **{format_duration(per_student)}**"
        )

        # ── 数据表 ──
        if has_id:
            st.dataframe(
                [{"序号": r[0], "学号": r[1], "姓名": r[2], "开始": r[3], "结束": r[4]} for r in rows],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.dataframe(
                [{"序号": r[0], "姓名": r[2], "开始": r[3], "结束": r[4]} for r in rows],
                use_container_width=True,
                hide_index=True,
            )

        # ── 下载 Excel ──
        excel_data = to_excel(rows, per_student)
        label = "📥 下载 Excel 答辩顺序表（含学号）" if has_id else "📥 下载 Excel 答辩顺序表"
        fname = f"答辩顺序_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label=label,
            data=excel_data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
