# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

RainClassroomAttendanceAnalyzer — Streamlit-based attendance analysis tool for 长江雨课堂 (changjiang.yuketang.cn). Upload Rain Classroom Excel exports and generate color-coded attendance sheets with per-session statistics.

## Commands

- **Run Streamlit web app**: `python -m streamlit run app.py`
- **Run GUI desktop app**: `python attendance_gui.py`
- **Build exe**: `pip install pyinstaller && python build_exe.py`
- **Run offline CLI script**: `python attendance_analyzer.py <input.xlsx>`
- **Dependencies**: `pip install streamlit openpyxl pandas pyinstaller`

## Project Structure

```
RainClassroomAttendanceAnalyzer/
├── app.py                    # Streamlit web app
├── attendance_gui.py         # tkinter GUI desktop app
├── attendance_analyzer.py    # CLI version of the parser
├── build_exe.py              # PyInstaller build script
├── 测试表格.xlsx             # Demo data file
├── dist/                     # Built exe output (gitignored)
├── requirements.txt
├── README.md
└── CLAUDE.md
```

## Architecture

### Data flow
1. User uploads Rain Classroom Excel (or clicks "加载示例数据")
2. `parse_summary()` reads the first sheet — maps 签到方式/得分 columns to session headers using nearest-header-left algorithm
3. `parse_sub_sheets()` reads sheets containing "课堂情况" in name — extracts 病假/事假 remarks from column G
4. Attendance status logic: known values (扫二维码/"正在上课"提示/教师添加/课堂暗号) → 上课; 未上课 + leave data → 病假/事假; 未上课 + no leave → 旷课
5. Output: 考勤明细 sheet (color-coded) + 课堂表现 sheet (top 10% gold highlight)

### Key files
- `app.py` — Streamlit UI + all parsing functions in one file
- `attendance_analyzer.py` — Same parsing logic, CLI interface, standalone use
- `测试表格.xlsx` — Synthetic demo file with 10 students, 3 sessions

## Important Notes

- Column mapping uses `hi <= ci` (nearest session header to the left, inclusive) — critical for lab sessions where sign column is directly under the session header
- Session labels come from Row 1 of the summary sheet, used as-is (no regex extraction)
- Sub-sheets are matched sequentially by order in workbook, not by name
- The present_set in both files must use Chinese quotation marks `“”` around "正在上课" — ASCII quotes `"` will not match
