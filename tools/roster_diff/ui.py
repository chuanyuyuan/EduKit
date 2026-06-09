"""
名单比对 — Streamlit UI 组件
参照 setdiff.com 风格：两个文本域 + 差异结果三栏展示。
"""
import streamlit as st
from .core import diff_sets

SAMPLE_A = """赵一一
钱二二
孙三三
李四四
周五五
吴六六
郑七七
冯九九
陈十十"""

SAMPLE_B = """赵一一
钱二二
孙三三
李四四
周五五
吴六六
郑七七
王八八"""


def _clear():
    st.session_state.roster_a = ""
    st.session_state.roster_b = ""
    st.session_state.roster_diff_result = None


def _display_result(result):
    stats = result['stats']

    st.markdown("---")
    meta_cols = st.columns(4)
    meta_cols[0].metric("列表 A", f"{stats['a_count']} 条")
    meta_cols[1].metric("列表 B", f"{stats['b_count']} 条")
    meta_cols[2].metric("交集 (A∩B)", f"{stats['both_count']} 条")
    meta_cols[3].metric("仅 A / 仅 B",
                        f"{stats['only_a_count']} 条 / {stats['only_b_count']} 条")

    if stats['both_count'] == stats['a_count'] == stats['b_count']:
        st.success("两份名单完全一致，无差异。")
        return

    tab_both, tab_only_a, tab_only_b = st.tabs([
        f":material/select_all: 交集 ({stats['both_count']})",
        f":material/flag: 仅列表 A ({stats['only_a_count']})",
        f":material/outlined_flag: 仅列表 B ({stats['only_b_count']})",
    ])

    with tab_both:
        if result['both']:
            st.code("\n".join(result['both']), language="text")
        else:
            st.info("无交集")

    with tab_only_a:
        if result['only_a']:
            st.code("\n".join(result['only_a']), language="text")
        else:
            st.info("无差异")

    with tab_only_b:
        if result['only_b']:
            st.code("\n".join(result['only_b']), language="text")
        else:
            st.info("无差异")


def render_page():
    st.header(":material/compare_arrows: setDiff工具")
    st.markdown("""
在两边的文本框中粘贴内容，快速找出两个集合的差异。

**支持的内容：**
- 学生名单、学号、工号等文本数据
- 每行视为一个独立个体

**操作步骤：**
1. 在左侧文本框粘贴列表 A
2. 在右侧文本框粘贴列表 B
3. 点击 **开始比对** 查看差异

**功能特点：**
- 自动识别交集、仅 A 有、仅 B 有三类数据
- 结果区域自带复制按钮（鼠标悬停可见）
- 两份名单完全一致时有明确提示
""")

    if st.button(":material/lightbulb: 加载示例数据看看效果", use_container_width=True):
        st.session_state.roster_a = SAMPLE_A
        st.session_state.roster_b = SAMPLE_B
        st.session_state.roster_diff_result = diff_sets(SAMPLE_A.strip(), SAMPLE_B.strip())

    col_left, col_right = st.columns(2)
    with col_left:
        st.text_area("列表 A", height=250, key="roster_a",
                     placeholder="张三\n李四\n王五\n每行一个独立个体")
    with col_right:
        st.text_area("列表 B", height=250, key="roster_b",
                     placeholder="张三\n王五\n赵六\n每行一个独立个体")

    b1, b2, _ = st.columns([1, 1, 6])
    with b1:
        clicked = st.button(":material/play_arrow: 开始比对", type="primary", use_container_width=True)

    ta = st.session_state.roster_a.strip()
    tb = st.session_state.roster_b.strip()

    if clicked:
        if ta and tb:
            st.session_state.roster_diff_result = diff_sets(ta, tb)
        else:
            st.session_state.roster_diff_result = None

    result = st.session_state.get("roster_diff_result")
    if not result or not ta or not tb:
        if not ta and not tb:
            st.info("请在上方两侧输入要比较的名单。")
        elif not result:
            st.info("点击「开始比对」查看差异。")
        return

    with b2:
        st.button(":material/clear: 清除", on_click=_clear, use_container_width=True)

    _display_result(result)
