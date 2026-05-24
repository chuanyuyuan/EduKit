# RainClassroomAttendanceAnalyzer

基于 Streamlit 的[长江雨课堂](https://changjiang.yuketang.cn/web/?index)考勤数据分析工具。上传雨课堂导出的 Excel 文件，自动生成带颜色标注的考勤明细和课堂表现统计。

## 功能

- 解析雨课堂汇总表，自动映射签到方式和得分列到对应课堂
- 读取子表提取病假/事假信息
- 生成考勤明细（上课/旷课/病假/事假）并标注颜色
- 统计无故旷课率和总旷课率
- 生成课堂表现得分表，前 10% 学生高亮
- 支持在线预览和下载 Excel

## 快速开始

```bash
pip install streamlit openpyxl pandas
streamlit run app.py
```

## 部署

将仓库连接到 [Streamlit Community Cloud](https://streamlit.io/cloud) 即可一键部署。

## 目录结构

```
RainClassroomAttendanceAnalyzer/
├── app.py                    # Streamlit 在线工具
├── attendance_analyzer.py    # 本地版分析脚本
├── 测试表格.xlsx             # 示例数据
└── CLAUDE.md
```

## 技术栈

- **Streamlit** — 前端界面
- **openpyxl** — Excel 读写
- **pandas** — 数据展示
