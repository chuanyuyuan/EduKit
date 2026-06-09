# EduKit 教师工具包

基于 Streamlit 的教师常用工具集合，涵盖长江雨课堂考勤分析、实验报告查重、名单比对。提供 Web 在线版（Streamlit，含全部工具）、桌面版（tkinter GUI，仅考勤分析）和 CLI 命令行版（仅考勤分析）三种使用方式。

## 工具

### 雨课堂数据分析

解析长江雨课堂导出的 Excel 考勤表，自动生成带颜色标注的考勤明细和过程性成绩记载表。

- 自动识别文件类型（批量导出汇总表 / 单次课导出）
- 单文件分析和两表合并（如理论班 + 实验班），自动校验学生名单一致性
- **考勤明细表** — 每次课的考勤状态（上课/旷课/病假/事假），颜色标注，含无故旷课率和总旷课率
- **课堂表现得分** — 各次课得分汇总，前 10% 学生高亮
- **过程性成绩记载表** — ✓/✗/△ 符号标记，宋体 9pt、居中、细边框

### 头歌图片查重

解压学生提交的 ZIP 压缩包，提取 Word 文档中的嵌入图片，通过像素级 MD5 指纹交叉比对检测抄袭。

- 自动解压并整理学生实验报告，跳过「答题记录」
- 提取嵌入图片像素 MD5 指纹（消除 WPS Writer / Word 元数据干扰）
- 共享 ≥ 3 张图片判定为疑似抄袭
- 支持 .doc（OLE2 流扫描）和 .docx 格式
- 生成带饼图 + 颜色标注的 Excel 报表

### setDiff 工具

快速比对两份名单的差异，自动去重并显示交集和差集。

- 支持大小写忽略模式
- 自动 trim 空白和去重
- 适用于查未交作业、查缺勤等场景

## 快速开始

### 安装依赖

```bash
pip install streamlit openpyxl pandas python-docx olefile Pillow pyinstaller
```

### Web 在线版（推荐）

```bash
python -m streamlit run app.py
```

### 桌面版（仅考勤分析）

```bash
python attendance_gui.py
```

### CLI 命令行版（仅考勤分析）

```bash
# 单文件模式
python attendance_analyzer.py <文件.xlsx>

# 合并模式（两文件）
python attendance_analyzer.py <文件一.xlsx> <文件二.xlsx>
```

### 打包为 exe

```bash
python build_exe.py
```

exe 生成在 `dist/` 目录，双击即可运行，无需 Python 环境。

## 项目结构

```
RainClassroomAttendanceAnalyzer/
├── app.py                    # Streamlit Web 应用（含首页导航）
├── attendance_gui.py         # tkinter 桌面版
├── attendance_analyzer.py    # CLI 命令行版
├── build_exe.py              # PyInstaller 打包脚本
├── samples/
│   ├── sample_attendance.xlsx      # 考勤分析示例数据
│   └── sample_report_checker.zip   # 查重分析示例数据
├── tools/
│   ├── attendance/           # 雨课堂考勤分析
│   │   ├── core.py           #   核心解析逻辑
│   │   ├── ui.py             #   Streamlit UI
│   │   └── tests/
│   │       └── test_ui.py    #    Playwright UI 测试
│   ├── roster_diff/          # 名单比对
│   │   ├── core.py           #   核心逻辑
│   │   ├── ui.py             #   Streamlit UI
│   │   └── tests/
│   │       ├── test_core.py  #   核心逻辑测试
│   │       └── test_ui.py    #    Playwright UI 测试
│   └── report_checker/       # 头歌图片查重
│       ├── core.py           #   流水线：解压→整理→指纹提取→比对→报表
│       ├── ui.py             #   Streamlit UI
│       └── tests/
│           ├── test_core.py  #   核心逻辑测试
│           ├── test_real_data.py  # 合成数据 e2e 测试
│           └── test_ui.py    #    Playwright UI 测试
├── tests/
│   ├── test_analyzer.py      # CLI 版功能测试（121 项）
│   ├── test_app.py           # Streamlit 版功能测试（34 项）
│   ├── test_gui.py           # GUI 版功能测试（95 项）
│   ├── sample_single_1.xlsx  # 测试数据
│   ├── sample_single_2.xlsx
│   ├── sample_summary_1.xlsx
│   └── sample_summary_2.xlsx
├── requirements.txt
└── README.md
```

## 测试覆盖

| 测试 | 数量 | 运行命令 |
|------|------|---------|
| 考勤 CLI 版 | 121 | `python tests/test_analyzer.py` |
| 考勤 Web 版 | 34 | `python tests/test_app.py` |
| 考勤 GUI 版 | 95 | `python tests/test_gui.py` |
| 名单比对核心 | 59 | `python tools/roster_diff/tests/test_core.py` |
| 名单比对 UI | 15 | `python -m tools.roster_diff.tests.test_ui` |
| 查重核心 | 58 | `python tools/report_checker/tests/test_core.py` |
| 查重合成数据 e2e | 53 | `python tools/report_checker/tests/test_real_data.py` |
| 查重 UI | 2 | `python -m tools.report_checker.tests.test_ui` |
| 考勤 UI | 4 | `python -m tools.attendance.tests.test_ui` |
| **合计** | **441** | |

运行全部测试：

```bash
python tests/test_analyzer.py && python tests/test_app.py && python tests/test_gui.py && python tools/roster_diff/tests/test_core.py && python tools/report_checker/tests/test_core.py && python tools/report_checker/tests/test_real_data.py
```

## 技术栈

- **Streamlit** — Web 界面及多工具导航
- **tkinter** — 桌面 GUI
- **openpyxl** — Excel 读写与报表生成
- **python-docx** — Word 文档图片提取
- **Pillow** — 图片像素级解码与指纹计算
- **olefile** — .doc 格式 OLE2 流解析
- **Playwright** — UI 自动化测试
- **PyInstaller** — exe 打包

## CI/CD

推送 `v*` 标签到 GitHub 时自动触发：

```bash
git tag v1.1.0
git push origin --tags
```

自动流程：运行测试 → 打包 exe → 创建 Release 并上传附件。
