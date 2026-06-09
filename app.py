#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
EduKit — 教师工具包
"""
import streamlit as st

st.set_page_config(page_title="EduKit 教师工具包", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 2.5rem; padding-bottom: 3rem; }
    .material-symbols-outlined { vertical-align: middle; margin-right: 0.1em; }
    [data-testid="column"] { display: flex; }
    .card { border:1px solid #ddd; border-radius:12px; padding:1.8rem 1.5rem;
            width:100%; height:100%; display:flex; flex-direction:column;
            justify-content:space-between; box-sizing:border-box; }
    .card:hover { border-color:#aaa; box-shadow:0 2px 12px rgba(0,0,0,.08); }
    section[data-testid="stSidebar"] a[href="/"] {
        font-weight:700; font-size:1.1rem; border-bottom:1px solid #eee;
        padding-bottom:0.35rem; margin-bottom:0.25rem;
    }
</style>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" />
""", unsafe_allow_html=True)


from tools.attendance.ui import render_attendance_page
from tools.roster_diff.ui import render_page as roster_diff_page_fn
from tools.report_checker.ui import render_page as report_checker_page_fn
from tools.learning_analytics.ui import render_page as learning_analytics_page_fn

attendance_page = st.Page(render_attendance_page, title="雨课堂课堂数据分析", icon=":material/calendar_month:", url_path="attendance")
roster_diff_page = st.Page(roster_diff_page_fn, title="setDiff工具", icon=":material/compare_arrows:", url_path="roster_diff")
report_checker_page = st.Page(report_checker_page_fn, title="头歌图文实验图片查重", icon=":material/file_copy:", url_path="report_checker")
learning_analytics_page = st.Page(learning_analytics_page_fn, title="雨课堂学情分析", icon=":material/insights:", url_path="learning_analytics")

def _landing_page():
    st.markdown("<h1 style='margin-bottom:1.5rem;'>EduKit 教师工具包</h1>", unsafe_allow_html=True)
    cols = st.columns(4, gap="small")

    cards = [
        ("attendance", "calendar_month", "雨课堂数据分析",
         "解析雨课堂 Excel 考勤表，生成带颜色标注的考勤明细和过程性成绩记载表。"
         "支持单次课和合并模式。"),
        ("learning_analytics", "insights", "雨课堂学情分析",
         "上传雨课堂数据，自动分析出勤与得分趋势，生成 AI 学情评语和学生画像。"),
        ("report_checker", "file_copy", "头歌图片查重",
         "解压学生 ZIP 压缩包，提取 Word 嵌入图片，通过像素级 MD5 指纹交叉比对检测抄袭。"),
        ("roster_diff", "compare_arrows", "setDiff工具",
         "快速比对两份名单差异，自动去重并显示交集和差集。支持大小写忽略。"),
    ]

    for col, (key, icon, title, desc) in zip(cols, cards):
        with col:
            st.markdown(
                f"<div class='card'>"
                f"<div><div style='font-size:2rem;margin-bottom:0.8rem;'>"
                f"<span class='material-symbols-outlined' style='font-size:2rem;'>{icon}</span></div>"
                f"<h3 style='margin-top:0;'>{title}</h3>"
                f"<p style='color:#555;line-height:1.6;flex-grow:1;'>{desc}</p></div>"
                f"<div style='margin-top:1.2rem;'>"
                f"<a href='/{key}' target='_self' style='"
                f"display:inline-block;padding:0.5rem 1.2rem;border-radius:8px;"
                f"background:#0d6efd;color:#fff;text-decoration:none;font-size:0.9rem;"
                f"text-align:center;'>"
                f"打开工具 →</a></div></div>",
                unsafe_allow_html=True,
            )

root_page = st.Page(_landing_page, title="EduKit")

with st.sidebar:
    st.page_link(root_page, label="EduKit 教师工具包", use_container_width=True)
    st.page_link(attendance_page, label="雨课堂课堂数据分析", icon=":material/calendar_month:", use_container_width=True)
    st.page_link(learning_analytics_page, label="雨课堂学情分析", icon=":material/insights:", use_container_width=True)
    st.page_link(report_checker_page, label="头歌图片查重", icon=":material/file_copy:", use_container_width=True)
    st.page_link(roster_diff_page, label="setDiff工具", icon=":material/compare_arrows:", use_container_width=True)

    st.divider()

pg = st.navigation([root_page, attendance_page, learning_analytics_page, report_checker_page, roster_diff_page], position="hidden")
pg.run()
