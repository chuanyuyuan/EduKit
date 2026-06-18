"""
查重分析 — Streamlit UI 组件
"""
import os
import base64
import tempfile
import shutil

import streamlit as st
import pandas as pd

from .core import run_pipeline

DEMO_ZIP = 'samples/sample_report_checker.zip'


def _demo_zip_path():
    return os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        DEMO_ZIP,
    )


def _sample_zip_link():
    """Return base64 data URI for the demo ZIP file download link."""
    path = _demo_zip_path()
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f"data:application/zip;base64,{b64}"


def _sorted_results(results):
    """按判定结论排序：无图片优先，其余保持字母序。"""
    none = [r for r in results if r[1] == "无图片"]
    plag = [r for r in results if r[1] == "疑似抄袭"]
    normal = [r for r in results if r[1] == "正常"]
    return none + plag + normal


def _render_network(result):
    """力导向网络图：节点=学生，边按抄袭比例显示粗细。"""
    pairs = result.get("pairs", [])
    if not pairs:
        st.info("未检测到抄袭关联，无需展示关系网络。")
        return

    import networkx as nx
    from matplotlib import font_manager
    import matplotlib.pyplot as plt
    from matplotlib import rcParams

    _cn_candidates = [f.name for f in font_manager.fontManager.ttflist
                      if any(k in f.name.lower() for k in ("yahei", "simhei", "simsun", "noto", "wenquanyi",
                                                           "dengxian", "stxihei", "stsong",
                                                           "arial unicode"))]
    if not _cn_candidates:
        import os
        _fallback_fonts = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyhbd.ttc",
            "C:/Windows/Fonts/simsun.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        ]
        for _p in _fallback_fonts:
            if os.path.exists(_p):
                font_manager.fontManager.addfont(_p)
                _cn_candidates = [font_manager.FontProperties(fname=_p).get_name()]
                break
    rcParams['font.sans-serif'] = _cn_candidates + ['DejaVu Sans', 'sans-serif']
    rcParams['axes.unicode_minus'] = False

    img_counts = result.get("img_counts", {})
    verdicts = {r[0]: r[1] for r in result["results"]}
    G = nx.Graph()

    for name in verdicts:
        G.add_node(name)
    for n1, n2, c in pairs:
        total = min(img_counts.get(n1, 1), img_counts.get(n2, 1))
        ratio = c / total if total > 0 else 0
        G.add_edge(n1, n2, weight=c, ratio=ratio)

    if G.number_of_edges() == 0:
        st.info("无关联数据。")
        return

    pos = nx.spring_layout(G, k=3, iterations=100, seed=42)

    color_map = {"疑似抄袭": "#E74C3C", "正常": "#27AE60", "无图片": "#F39C12"}
    node_colors = [color_map.get(verdicts.get(n, "#95A5A6")) for n in G.nodes()]

    node_sizes = []
    for n in G.nodes():
        if G.degree(n) > 0:
            node_sizes.append(500 + 300 * G.degree(n))
        else:
            node_sizes.append(500)

    labels = {n: n for n in G.nodes()}

    edge_widths = [G[u][v]["ratio"] * 5 for u, v in G.edges()]
    edge_colors = [G[u][v]["ratio"] for u, v in G.edges()]

    fig, ax = plt.subplots(figsize=(14, 10))
    nx.draw_networkx_edges(G, pos, width=edge_widths, edge_color=edge_colors,
                           edge_cmap=plt.cm.Reds, alpha=0.6, ax=ax)
    nx.draw_networkx_nodes(G, pos, node_color=node_colors,
                           node_size=node_sizes, alpha=0.9, ax=ax)

    nx.draw_networkx_labels(G, pos, labels=labels, font_size=7,
                            font_color="#333333", ax=ax)

    ax.axis("off")
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    with st.expander("图例说明", expanded=True):
        st.markdown(f"""
- **红色节点** — 疑似抄袭（共 {result["plag_count"]} 人）
- **绿色节点** — 正常（共 {result["total"] - result["plag_count"] - result["none_count"]} 人）
- **橙色节点** — 无图片（共 {result["none_count"]} 人）
- **连线越粗/越红** — 抄袭比例越高（相同图片数/本人图片数）
- **节点越大** — 关联人数越多
""")


def _run_analysis(zip_path, workspace, status):
    """Run the full pipeline and return result dict."""
    log = lambda msg: status.write(msg)
    progress_bar = None

    def on_progress(current, total, msg):
        nonlocal progress_bar
        if progress_bar is None:
            progress_bar = st.progress(0, text=msg)
        progress_bar.progress(min(current / total, 1.0), text=msg)

    try:
        result = run_pipeline(
            zip_path, workspace,
            log_func=log,
            progress_callback=on_progress,
        )
    except Exception as e:
        st.error(f"分析过程出错。请确认上传的是头歌图文实验平台导出的有效 ZIP 压缩包。\n\n错误详情：{e}")
        return None

    if progress_bar is not None:
        progress_bar.empty()

    if result and result["total"] > 0:
        status.update(
            label=f"分析完成：共 {result['total']} 份报告",
            state="complete", expanded=False)
    else:
        status.update(label="分析完成", state="complete", expanded=False)

    return result


def render_page():
    st.header(":material/file_copy: 头歌图文实验图片查重")
    link = _sample_zip_link()
    st.markdown(f"""
上传学生从头歌图文实验平台导出的附件压缩包（ZIP），检测实验报告中是否存在图片复用/抄袭。

**适用平台：** [头歌平台图文实验](https://www.educoder.net/)

**支持的文件类型：**
- 平台导出的附件压缩包（ZIP 格式，内含学生 Word 实验报告）

**原始文件获取方式：**
1. 进入头歌图文实验空间，找到目标实验
2. 点击 **导出** → **导出答题记录与附件**
3. 下载生成的压缩包（即 `.zip` 格式的附件压缩包）。<a href="{link}" download="{DEMO_ZIP}">下载示例压缩包</a>

**上传后：**
- 自动解压、提取并整理所有学生 Word 实验报告
- 提取报告中嵌入的图片并计算指纹，交叉比对
- 相同图片数 ≥ 本人图片数 40% 判定为疑似抄袭
- 生成带颜色标注和饼图的查重报告 Excel
""", unsafe_allow_html=True)

    uploaded = st.file_uploader(":material/folder_zip: 选择 ZIP 压缩包")

    # Detect new upload (different file or re-upload) → clear result
    if uploaded:
        last = st.session_state.get("_rc_last_name")
        if last != uploaded.name:
            st.session_state.report_checker_result = None
            st.session_state._rc_last_name = uploaded.name
        st.session_state.rc_show_demo = False

    # Demo button
    if not uploaded:
        if st.button(":material/lightbulb: 加载示例数据看看效果", use_container_width=True):
            st.session_state.rc_show_demo = True
            st.rerun()

    result = st.session_state.get("report_checker_result")

    # Auto-run when source is ready but no result yet
    if result is None:
        should_run = False
        workspace = None
        zip_path = None

        if uploaded:
            if not uploaded.name.endswith('.zip'):
                st.error(f'不支持的文件格式："{uploaded.name}"，请上传 .zip 压缩包。')
            else:
                should_run = True
                workspace = tempfile.mkdtemp()
                zip_path = os.path.join(workspace, "upload.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded.getvalue())
        elif st.session_state.get("rc_show_demo"):
            demo_path = _demo_zip_path()
            if os.path.exists(demo_path):
                should_run = True
                workspace = tempfile.mkdtemp()
                zip_path = os.path.join(workspace, "demo.zip")
                with open(zip_path, "wb") as f:
                    with open(demo_path, "rb") as src:
                        f.write(src.read())

        if should_run:
            with st.status("正在分析...", expanded=True) as status:
                result = _run_analysis(zip_path, workspace, status)
            shutil.rmtree(workspace, ignore_errors=True)

            if result:
                st.session_state.report_checker_result = result
                st.rerun()

    result = st.session_state.get("report_checker_result")

    if not result:
        if not uploaded and not st.session_state.get("rc_show_demo"):
            st.info("请上传 ZIP 文件开始分析。")
        return

    # ════════════════════════════════════════
    #  展示结果
    # ════════════════════════════════════════
    if st.session_state.get("rc_show_demo"):
        st.info("当前展示的是示例数据，你可以上传自己的 ZIP 文件替换。")
    st.markdown("---")

    normal_count = result["total"] - result["plag_count"] - result["none_count"]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric(":material/description: 总报告数", f"{result['total']} 份")
    m2.metric(":material/report: 疑似抄袭", f"{result['plag_count']} 份")
    m3.metric(":material/image: 无图片异常", f"{result['none_count']} 份")
    m4.metric(":material/check_circle: 正常", f"{normal_count} 份")

    tab1, tab2, tab3 = st.tabs([":material/format_list_bulleted: 详细名单", ":material/pie_chart: 统计概览", ":material/hub: 关系网络"])

    with tab1:
        if result["results"]:
            sorted_ = _sorted_results(result["results"])
            df = pd.DataFrame(sorted_, columns=["学号姓名", "判定结论", "详细依据"])
            styled = df.style.map(
                lambda v: "background-color: #FFC7CE" if v == "疑似抄袭"
                else "background-color: #FFEB9C" if v == "无图片"
                else ""
            )
            st.dataframe(styled, use_container_width=True)
        else:
            st.info("无有效报告数据。")

    with tab2:
        chart_data = pd.DataFrame({
            "分类": ["疑似抄袭", "无图异常", "正常报告"],
            "人数": [result["plag_count"], result["none_count"], normal_count],
        })
        chart_data = chart_data[chart_data["人数"] > 0]
        if not chart_data.empty:
            try:
                import altair as alt
                chart = alt.Chart(chart_data).mark_arc(innerRadius=30).encode(
                    theta=alt.Theta(field="人数", type="quantitative"),
                    color=alt.Color(
                        field="分类", type="nominal",
                        scale=alt.Scale(
                            domain=["疑似抄袭", "无图异常", "正常报告"],
                            range=["#E74C3C", "#F39C12", "#27AE60"],
                        ),
                    ),
                    tooltip=["分类", "人数"],
                ).properties(height=300)
                st.altair_chart(chart, use_container_width=True)
            except ImportError:
                st.dataframe(chart_data, hide_index=True)
        else:
            st.info("无统计数据。")

        st.subheader("分类统计")
        st.dataframe(chart_data, hide_index=True)

    with tab3:
        _render_network(result)

    st.divider()
    result["excel_buf"].seek(0)
    st.download_button(
        label=":material/download: 下载查重报告 Excel",
        data=result["excel_buf"],
        file_name="查重结果汇总报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
