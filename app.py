#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
雨课堂考勤分析 - Streamlit 在线工具
"""

import sys
import os
from collections import OrderedDict
from io import BytesIO

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="雨课堂考勤分析", layout="wide")

# ── 核心解析函数（与本地版一致） ──


def col_idx(col):
    i = 0
    for ch in col:
        i = i * 26 + (ord(ch) - 64)
    return i - 1


def parse_summary(wb):
    ws = wb[wb.sheetnames[0]]

    row2 = {}
    for cell in ws[2]:
        if cell.value:
            row2[cell.column_letter] = str(cell.value)

    row1 = {}
    for cell in ws[1]:
        if cell.value:
            row1[cell.column_letter] = str(cell.value)

    sign_cols = sorted(
        [c for c, v in row2.items() if v == '签到方式'],
        key=col_idx
    )
    score_cols = sorted(
        [c for c, v in row2.items() if v.startswith('得分')],
        key=col_idx
    )
    session_headers = sorted(row1.keys(), key=col_idx)

    def map_cols_to_sessions(cols):
        result = OrderedDict()
        keys = []
        for c in cols:
            ci = col_idx(c)
            nearest = None
            nearest_i = -1
            for sh in session_headers:
                hi = col_idx(sh)
                if hi <= ci and hi > nearest_i:
                    nearest = row1[sh]
                    nearest_i = hi
            if nearest:
                result[nearest] = c
                keys.append(nearest)
        return result, keys

    session_sign_map, session_keys = map_cols_to_sessions(sign_cols)
    session_score_map, _ = map_cols_to_sessions(score_cols)

    students = []
    for row in ws.iter_rows(min_row=3, values_only=False):
        cells = {}
        for c in row:
            if c.value is None:
                continue
            try:
                cells[c.column_letter] = str(c.value)
            except AttributeError:
                pass
        sid = cells.get('A', '')
        name = cells.get('D', '')
        if not sid and not name:
            continue

        rec = {
            'id': sid,
            'name': name,
            'attendance': {},
            'scores': {},
        }
        for sk, sc in session_sign_map.items():
            rec['attendance'][sk] = cells.get(sc, '')
        for sk, sc in session_score_map.items():
            raw = cells.get(sc, '')
            try:
                rec['scores'][sk] = float(raw) if raw else None
            except ValueError:
                rec['scores'][sk] = None
        students.append(rec)

    return session_keys, session_sign_map, session_score_map, students


def parse_sub_sheets(wb, session_keys):
    leave_data = {}
    idx = 0

    for name in wb.sheetnames:
        if '课堂情况' not in name:
            continue
        if idx >= len(session_keys):
            break

        session_key = session_keys[idx]
        idx += 1
        ws = wb[name]

        for row in ws.iter_rows(min_row=4, values_only=False):
            cells = {}
            for c in row:
                if c.value is None:
                    continue
                try:
                    cells[c.column_letter] = str(c.value)
                except AttributeError:
                    pass
            sid = cells.get('A', '')
            remark = cells.get('G', '')
            if sid and remark in ('病假', '事假'):
                leave_data[(session_key, sid)] = remark

    return leave_data


def generate_output(session_keys, session_sign_map, session_score_map, students, leave_data):
    """Generate output Excel bytes + summary data from parsed data."""

    present_set = {'扫二维码', '“正在上课”提示', '教师添加', '课堂暗号'}
    session_count = len(session_sign_map)

    # ── Build data for Sheet 1 ──
    out_rows = []
    for s in students:
        row = [s['id'], s['name']]
        absent_cnt = 0
        excused_cnt = 0
        for sk in session_sign_map:
            raw = s['attendance'].get(sk, '')
            if raw in present_set:
                status = '上课'
            elif raw == '未上课':
                leave = leave_data.get((sk, s['id']))
                if leave == '病假':
                    status = '病假'
                    excused_cnt += 1
                elif leave == '事假':
                    status = '事假'
                    excused_cnt += 1
                else:
                    status = '旷课'
                    absent_cnt += 1
            else:
                status = raw
            row.append(status)
        row.append(f'{absent_cnt / session_count:.0%}')
        row.append(f'{(absent_cnt + excused_cnt) / session_count:.0%}')
        out_rows.append(row)

    # ── Write output workbook ──
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    color_map = {'上课': green_fill, '病假': yellow_fill, '事假': yellow_fill, '旷课': red_fill}
    gold_fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')

    owb = Workbook()
    ows = owb.active
    ows.title = '考勤明细'
    headers = ['学号', '姓名'] + list(session_sign_map.keys()) + ['无故旷课率', '总旷课率']
    ows.append(headers)
    for row in out_rows:
        ows.append(row)
    for row in ows.iter_rows(min_row=2, max_row=ows.max_row,
                             min_col=3, max_col=2 + session_count):
        for cell in row:
            fill = color_map.get(cell.value)
            if fill:
                cell.fill = fill

    # ── Sheet 2: 课堂表现 ──
    ows2 = owb.create_sheet('课堂表现')
    score_headers = ['学号', '姓名'] + [k for k in session_score_map.keys()] + ['总分']
    ows2.append(score_headers)
    totals = []
    for s in students:
        row = [s['id'], s['name']]
        total = 0
        for sk in session_score_map:
            v = s['scores'].get(sk)
            if v is not None:
                row.append(v)
                total += v
            else:
                row.append('')
        row.append(total)
        totals.append(total)
        ows2.append(row)

    # Top 10% gold highlight
    sorted_totals = sorted(totals, reverse=True)
    n_top = max(1, round(len(sorted_totals) * 0.1))
    threshold = sorted_totals[n_top - 1] if n_top <= len(sorted_totals) else sorted_totals[-1]
    for row in ows2.iter_rows(min_row=2, max_row=ows2.max_row):
        if row[-1].value is not None and row[-1].value >= threshold:
            for cell in row:
                cell.fill = gold_fill

    buf = BytesIO()
    owb.save(buf)
    buf.seek(0)

    # ── Summary stats ──
    summary_lines = []
    for sk in session_sign_map:
        attend = absent = sick = personal = 0
        for s in students:
            raw = s['attendance'].get(sk, '')
            if raw in present_set:
                attend += 1
            elif raw == '未上课':
                leave = leave_data.get((sk, s['id']))
                if leave == '病假':
                    sick += 1
                elif leave == '事假':
                    personal += 1
                else:
                    absent += 1
        summary_lines.append({
            'session': sk,
            '上课': attend,
            '旷课': absent,
            '病假': sick,
            '事假': personal,
        })

    total_absent = sum(1 for s in students for sk in session_sign_map
                       if s['attendance'].get(sk) == '未上课'
                       and (sk, s['id']) not in leave_data)
    total_sick = sum(1 for v in leave_data.values() if v == '病假')
    total_personal = sum(1 for v in leave_data.values() if v == '事假')

    # ── Build processed data for preview ──
    session_names = list(session_sign_map.keys())
    preview_attendance = []
    for s in students:
        row = {'学号': s['id'], '姓名': s['name']}
        absent_cnt = 0
        excused_cnt = 0
        for sk in session_names:
            raw = s['attendance'].get(sk, '')
            if raw in present_set:
                status = '上课'
            elif raw == '未上课':
                leave = leave_data.get((sk, s['id']))
                if leave == '病假':
                    status = '病假'
                    excused_cnt += 1
                elif leave == '事假':
                    status = '事假'
                    excused_cnt += 1
                else:
                    status = '旷课'
                    absent_cnt += 1
            else:
                status = raw
            row[sk] = status
        row['无故旷课率'] = f'{absent_cnt / session_count:.0%}'
        row['总旷课率'] = f'{(absent_cnt + excused_cnt) / session_count:.0%}'
        preview_attendance.append(row)

    score_names = list(session_score_map.keys())
    preview_scores = []
    for s in students:
        row = {'学号': s['id'], '姓名': s['name']}
        total = 0
        for k in score_names:
            v = s['scores'].get(k)
            if v is not None:
                row[k] = v
                total += v
            else:
                row[k] = ''
        row['总分'] = total
        preview_scores.append(row)

    return buf, {
        'students': students,
        'session_count': session_count,
        'total_absent': total_absent,
        'total_sick': total_sick,
        'total_personal': total_personal,
        'summary_lines': summary_lines,
        'preview_headers': ['学号', '姓名'] + session_names + ['无故旷课率', '总旷课率'],
        'preview_attendance': preview_attendance,
        'score_headers': ['学号', '姓名'] + score_names + ['总分'],
        'preview_scores': preview_scores,
    }


# ── Streamlit UI ──

DEMO_FILE = '测试表格.xlsx'

st.title("长江雨课堂考勤数据分析工具")
st.markdown("上传雨课堂导出的 Excel 文件（汇总页 + 子表），自动生成考勤明细和课堂表现统计。")

st.markdown("""
**适用平台：** [长江雨课堂](https://changjiang.yuketang.cn/web/?index)

**原始文件获取方式：**
1. 登录长江雨课堂官网，进入你授课的班级
2. 点击 **批量导出数据**
3. 筛选需要统计的课堂记录
4. 下载导出的 Excel 文件（即 `.xlsx` 格式的汇总数据表）

**上传后：**
- 页面自动解析并展示考勤明细、课堂表现得分和统计摘要
- 点击下载按钮即可获取带颜色标注的完整 Excel 文件
""")

uploaded = st.file_uploader(
    "选择雨课堂导出的 Excel 文件",
    type=['xlsx'],
)

# ── Determine data source: uploaded file or demo ──
workbook_source = None
source_label = None
if uploaded:
    workbook_source = load_workbook(uploaded, data_only=True)
    source_label = "upload"
    st.session_state.show_demo = False
elif st.session_state.get("show_demo"):
    demo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DEMO_FILE)
    if os.path.exists(demo_path):
        workbook_source = load_workbook(demo_path, data_only=True)
        source_label = "demo"

if not uploaded and st.button("💡 加载示例数据看看效果", use_container_width=True):
    st.session_state.show_demo = True
    st.rerun()

# ── Process and show results ──
if workbook_source:
    if source_label == "demo":
        st.info("当前展示的是示例数据，你可以上传自己的文件替换。")

    status = st.status("正在解析...", expanded=True)
    with status:
        st.write("读取 Excel 文件...")
        wb = workbook_source

        st.write("解析汇总表（签到方式、得分列映射）...")
        session_keys, session_sign_map, session_score_map, students = parse_summary(wb)

        st.write("解析子表（病假/事假）...")
        leave_data = parse_sub_sheets(wb, session_keys)
        wb.close()

        st.write("生成输出文件...")
        buf, info = generate_output(session_keys, session_sign_map, session_score_map, students, leave_data)
    status.update(label="解析完成", state="complete", expanded=False)

    # ── 预览表格 ──
    st.subheader(f"考勤概况 — 共 {info['session_count']} 次课，{len(info['students'])} 名学生")

    tab1, tab2, tab3 = st.tabs(["考勤明细", "课堂表现", "统计摘要"])

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

        st.subheader("每次课统计")
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

    # ── 下载 ──
    st.divider()
    st.download_button(
        label="📥 下载考勤明细 Excel",
        data=buf,
        file_name="学生考勤明细表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("请上传文件开始分析。")
