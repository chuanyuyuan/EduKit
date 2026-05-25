# RainClassroomAttendanceAnalyzer

[长江雨课堂](https://changjiang.yuketang.cn/web/?index)考勤数据分析工具。支持在线版（Streamlit）和桌面版（tkinter GUI），可打包为 exe 免安装运行。

## 功能

- **自动识别文件** — 支持批量导出汇总表（含汇总页 + 课堂情况子表）和单次课导出文件
- **两种模式** — 单文件分析和两表合并（如理论班 + 实验班），自动校验学生名单一致性
- **考勤明细表** — 每次课的考勤状态（上课/旷课/病假/事假），颜色标注，含无故旷课率和总旷课率
- **课堂表现得分** — 各次课得分汇总，前 10% 学生高亮
- **过程性成绩记载表** — ✓/✗/△ 符号标记，宋体 9pt、居中、细边框，与样表格式一致
- **可离线运行** — 桌面版无需联网，可打包为 exe 免安装 Python

## 快速开始

### 在线版（Streamlit）

```bash
pip install streamlit openpyxl pandas
streamlit run app.py
```

部署到 [Streamlit Community Cloud](https://streamlit.io/cloud) 即可在线使用。

### 桌面版（tkinter GUI）

```bash
pip install openpyxl
python attendance_gui.py
```

### 打包为 exe

```bash
pip install pyinstaller
python build_exe.py
```

exe 生成在 `dist/` 目录，双击即可运行，无需 Python 环境。

> 已发布的 exe 可从 [Releases](https://github.com/chuanyuyuan/RainClassroomAttendanceAnalyzer/releases) 页面直接下载。

## 目录结构

```
RainClassroomAttendanceAnalyzer/
├── app.py                    # Streamlit 在线工具
├── attendance_gui.py         # tkinter 桌面版
├── attendance_analyzer.py    # CLI 分析脚本
├── build_exe.py              # PyInstaller 打包脚本
├── 示例表格.xlsx             # 示例数据
├── requirements.txt
├── CLAUDE.md
└── README.md
```

## 技术栈

- **Streamlit** — Web 界面
- **tkinter** — 桌面 GUI
- **openpyxl** — Excel 读写
- **PyInstaller** — exe 打包
