"""
雨课堂考勤分析 — Streamlit UI 组件
"""
import os
import base64
from io import BytesIO
from collections import OrderedDict

import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from .core import parse_file, generate_output, generate_process_score_sheet

DEMO_FILE = 'samples/sample_attendance.xlsx'


def _sample_link():
    """Return base64 data URI for the demo file download link."""
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), DEMO_FILE)
    if not os.path.exists(path):
        return ''
    with open(path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode()
    return f'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}'


def show_results(buf, info):
    """Display preview tables and download buttons."""
    st.subheader(f"考勤概况 — 共 {info['session_count']} 次课，{len(info['students'])} 名学生")

    tab1, tab2, tab3 = st.tabs([":material/table: 考勤明细", ":material/star: 课堂表现", ":material/bar_chart: 统计摘要"])

    with tab1:
        st.info("完整数据请下载 Excel 文件查看。下面预览前 10 名学生：")
        df = pd.DataFrame(info['preview_attendance'][:10])
        styled = df.style.map(
            lambda v: 'background-color: #C6EFCE' if v == '上课'
            else 'background-color: #FFEB9C' if v in ('病假', '事假')
            else 'background-color: #FFC7CE' if v == '旷课'
            else ''
        )
        st.dataframe(styled, use_container_width=True)

    with tab2:
        st.info("课堂表现得分详情请下载 Excel 文件查看。下面预览前 10 名学生：")
        all_scores = info['preview_scores']
        totals = [s['总分'] for s in all_scores]
        sorted_totals = sorted(totals, reverse=True)
        n_top = max(1, round(len(sorted_totals) * 0.1))
        threshold = sorted_totals[n_top - 1]

        df_scores = pd.DataFrame(all_scores[:10])
        styled_scores = df_scores.style.apply(
            lambda row: ['background-color: #FFD700'] * len(row)
            if row['总分'] >= threshold else [''] * len(row),
            axis=1,
        )
        st.dataframe(styled_scores, use_container_width=True)

    with tab3:
        col1, col2, col3 = st.columns(3)
        col1.metric("旷课总人次", info['total_absent'])
        col2.metric("病假总人次", info['total_sick'])
        col3.metric("事假总人次", info['total_personal'])

        st.subheader(":material/calendar_view_week: 每次课统计")
        summary_df = []
        for line in info['summary_lines']:
            summary_df.append({
                '课程': line['session'],
                '上课': line['上课'],
                '旷课': line['旷课'],
                '病假': line['病假'],
                '事假': line['事假'],
            })
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    st.divider()
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label=":material/download: 下载考勤明细 Excel",
            data=buf,
            file_name="学生考勤明细表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_dl2:
        st.download_button(
            label=":material/description: 下载过程性成绩记载表",
            data=info['process_score_buf'],
            file_name="过程性成绩记载表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def single_page():
    """单文件模式页面"""
    uploaded = st.file_uploader("选择雨课堂导出的 Excel 文件（仅支持 .xlsx 格式）")

    workbook_source = None
    source_label = None
    if uploaded:
        if not uploaded.name.endswith('.xlsx'):
            st.error(f'不支持的文件格式："{uploaded.name}"，请上传 .xlsx 文件。')
        else:
            try:
                workbook_source = load_workbook(uploaded, data_only=True)
                source_label = "upload"
                st.session_state.att_show_demo = False
            except Exception:
                st.error("无法读取该文件，请确认上传的是有效的 .xlsx 格式文件。")
    elif st.session_state.get("att_show_demo"):
        demo_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), DEMO_FILE)
        if os.path.exists(demo_path):
            try:
                workbook_source = load_workbook(demo_path, data_only=True)
                source_label = "demo"
            except Exception:
                st.error("无法读取示例文件。")
                st.session_state.att_show_demo = False

    if not uploaded and not st.session_state.get("att_show_demo"):
        if st.button(":material/lightbulb: 加载示例数据看看效果", use_container_width=True):
            st.session_state.att_show_demo = True
            st.rerun()

    if workbook_source:
        if source_label == "demo":
            st.info("当前展示的是示例数据，你可以上传自己的文件替换。")

        try:
            status = st.status("正在解析...", expanded=True)
            with status:
                st.write("读取 Excel 文件...")
                wb = workbook_source
                session_keys, session_sign_map, session_score_map, students, leave_data = parse_file(wb)
                wb.close()
                st.write("生成输出文件...")
                buf, info = generate_output(session_keys, session_sign_map, session_score_map, students, leave_data)
            status.update(label="解析完成", state="complete", expanded=False)
            show_results(buf, info)
        except ValueError as e:
            st.error(str(e))
        except Exception:
            st.error("解析过程出现未知错误，请确认上传的是雨课堂批量导出的汇总数据表（.xlsx）。")
    else:
        st.info("请上传文件开始分析。")


def merge_page():
    """合并模式页面"""
    st.caption("自动按学号匹配合并，支持学生顺序不一致。若两文件名单有差异会提示错误。")
    col1, col2 = st.columns(2)
    with col1:
        f1 = st.file_uploader("选择文件一", key="f1")
    with col2:
        f2 = st.file_uploader("选择文件二", key="f2")

    if f1 and f2:
        errors = []
        if not f1.name.endswith('.xlsx'):
            errors.append(f'文件一格式不支持："{f1.name}"')
        if not f2.name.endswith('.xlsx'):
            errors.append(f'文件二格式不支持："{f2.name}"')
        if errors:
            st.error('\n'.join(errors))
            st.stop()

        try:
            status = st.status("正在解析...", expanded=True)
            with status:
                st.write("解析文件一...")
                wb1 = load_workbook(f1, data_only=True)
                sk1, ssm1, sscm1, students1, ld1 = parse_file(wb1)
                wb1.close()

                st.write("解析文件二...")
                wb2 = load_workbook(f2, data_only=True)
                sk2, ssm2, sscm2, students2, ld2 = parse_file(wb2)
                wb2.close()

                st.write("验证学生信息一致性...")
                ids1 = {(s['id'], s['name']) for s in students1}
                ids2 = {(s['id'], s['name']) for s in students2}
                if ids1 != ids2:
                    diff = []
                    name1_by_id = {s['id']: s['name'] for s in students1}
                    name_mismatch_ids = set()
                    for s in students2:
                        if s['id'] in name1_by_id and name1_by_id[s['id']] != s['name']:
                            diff.append(f'  学号 {s["id"]} 姓名不一致：文件一为"{name1_by_id[s["id"]]}"，文件二为"{s["name"]}"')
                            name_mismatch_ids.add(s['id'])
                    for sid, sname in ids1 - ids2:
                        if sid not in name_mismatch_ids:
                            diff.append(f"  文件一有但文件二缺少：学号 {sid} {sname}")
                    for sid, sname in ids2 - ids1:
                        if sid not in name_mismatch_ids:
                            diff.append(f"  文件二有但文件一缺少：学号 {sid} {sname}")
                    raise ValueError("两个文件的学生名单不一致：\n" + "\n".join(diff))

                st.write("合并数据...")
                session_keys = sk1 + sk2
                session_sign_map = OrderedDict(list(ssm1.items()) + list(ssm2.items()))
                session_score_map = OrderedDict(list(sscm1.items()) + list(sscm2.items()))
                stus2_by_id = {s['id']: s for s in students2}
                students = []
                for s1 in students1:
                    s2 = stus2_by_id[s1['id']]
                    students.append({
                        'id': s1['id'],
                        'name': s1['name'],
                        'cls': s1['cls'],
                        'attendance': {**s1['attendance'], **s2['attendance']},
                        'scores': {**s1['scores'], **s2['scores']},
                    })
                leave_data = {**ld1, **ld2}

                st.write("生成输出文件...")
                buf, info = generate_output(session_keys, session_sign_map, session_score_map, students, leave_data)
            status.update(label="解析完成", state="complete", expanded=False)
            show_results(buf, info)
        except ValueError as e:
            st.error(str(e))
        except Exception:
            st.error("解析过程出现未知错误，请确认上传的是雨课堂批量导出的汇总数据表（.xlsx）。")
    else:
        st.info("请上传两个考勤文件开始合并分析。")


def render_attendance_page():
    """考勤分析工具入口 — 说明 + 模式选择 + 单文件/合并模式路由。"""
    st.header(":material/calendar_month: 雨课堂课堂数据分析")
    st.markdown(f"""
上传雨课堂导出的 Excel 文件，自动生成考勤明细和课堂表现统计。

**适用平台：** [长江雨课堂](https://changjiang.yuketang.cn/web/?index)

**支持的文件类型：**
- 批量导出的汇总数据表（含汇总页 + 课堂情况子表）
- 单次课导出的课堂数据表（自动识别）

**原始文件获取方式：**
1. 登录长江雨课堂官网，进入你授课的班级
2. 点击 **批量导出数据**
3. 筛选需要统计的课堂记录
4. 下载导出的 Excel 文件（即 `.xlsx` 格式的汇总数据表）。<a href="{_sample_link()}" download="{DEMO_FILE}">下载示例表格</a>

**上传后：**
- 自动解析并分 tab 展示考勤明细、课堂表现得分和统计摘要
- 支持单文件分析和两表合并（如理论班 + 实验班）
- 提供两种下载：考勤明细 Excel（颜色标注）和过程性成绩记载表（✓/✗/△）
""", unsafe_allow_html=True)

    st.markdown("""
    <style>
        button[kind='segmented_control'], button[kind='segmented_controlActive'] {
            font-size: 1.8rem !important;
            padding: 1rem 2.5rem !important;
        }
    </style>
    """, unsafe_allow_html=True)

    mode = st.segmented_control("模式", [":material/file_present: 单文件模式", ":material/merge: 合并模式"],
                                default=":material/file_present: 单文件模式", label_visibility="collapsed")

    if "单文件模式" in mode:
        single_page()
    else:
        merge_page()
