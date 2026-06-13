"""答辩顺序生成器 — Streamlit UI"""
import streamlit as st
from .core import parse_names, generate_schedule, format_duration, to_excel

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


def render_page():
    st.title("🎤 答辩顺序生成器")
    st.caption("输入学生姓名和答辩时间区间，随机排序后自动分配每人答辩时间。")

    names_text = st.text_area(
        "学生姓名",
        value=SAMPLE_NAMES,
        height=200,
        placeholder="每行一个姓名，或用逗号、顿号分隔",
    )

    col1, col2 = st.columns(2)
    with col1:
        start_time = st.text_input("开始时间", value="08:00", placeholder="08:00")
    with col2:
        end_time = st.text_input("结束时间", value="09:30", placeholder="09:30")

    if st.button("🎲 随机生成", type="primary", use_container_width=True):
        names = parse_names(names_text)
        if not names:
            st.error("请至少输入一个姓名。")
            return

        try:
            rows, per_student = generate_schedule(names, start_time, end_time)
        except ValueError as e:
            st.error(str(e))
            return

        st.success(f"共 {len(names)} 名学生，{start_time}–{end_time}，每人 **{format_duration(per_student)}**")

        st.dataframe(
            [{"序号": r[0], "姓名": r[1], "开始": r[2], "结束": r[3]} for r in rows],
            use_container_width=True,
            hide_index=True,
        )

        excel_data = to_excel(rows, per_student)

        import datetime as _dt
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 下载 Excel 答辩顺序表",
            data=excel_data,
            file_name=f"答辩顺序_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
