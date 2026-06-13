"""答辩顺序生成器 — 核心逻辑"""
import random
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def parse_names(text: str) -> list[str]:
    """从文本中提取学生姓名列表。支持换行、逗号、顿号分隔。"""
    import re
    names = re.split(r"[\n,，、；;\s]+", text.strip())
    return [n.strip() for n in names if n.strip()]


def generate_schedule(names: list[str], start_time, end_time):
    """
    随机打乱顺序，按时间区间分配每人时间。

    start_time, end_time 可以是 datetime.time 或 "HH:MM" 字符串。

    返回 (schedule_rows, per_student_seconds):
      schedule_rows: [(序号, 姓名, 开始_str, 结束_str), ...]
    """
    random.shuffle(names)
    n = len(names)
    if isinstance(start_time, str):
        start_time = datetime.strptime(start_time, "%H:%M").time()
    if isinstance(end_time, str):
        end_time = datetime.strptime(end_time, "%H:%M").time()
    t1 = datetime.combine(datetime.today(), start_time)
    t2 = datetime.combine(datetime.today(), end_time)
    if t2 <= t1:
        raise ValueError("结束时间必须晚于开始时间")
    total_seconds = int((t2 - t1).total_seconds())
    per_student = total_seconds / n

    rows = []
    for i, name in enumerate(names):
        start = t1 + timedelta(seconds=int(i * per_student))
        end = t1 + timedelta(seconds=int((i + 1) * per_student))
        rows.append((i + 1, name, start.strftime("%H:%M"), end.strftime("%H:%M")))

    return rows, per_student


def format_duration(seconds: float) -> str:
    m = int(seconds // 60)
    s = round(seconds % 60)
    if s == 0:
        return f"{m}分钟"
    return f"{m}分{s}秒"


def to_excel(rows: list, per_student_seconds: float) -> bytes:
    """生成 Excel 文件并返回二进制内容。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "答辩顺序表"

    header_font = Font(name="微软雅黑", bold=True, size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    cell_font = Font(name="微软雅黑", size=10.5)
    center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # 表头
    headers = ["答辩序号", "姓名", "开始时间", "结束时间"]
    widths = [12, 18, 14, 14]
    for ci, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w

    # 数据行
    for ri, (seq, name, t1, t2) in enumerate(rows, start=2):
        for ci, val in enumerate([seq, name, t1, t2], start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = cell_font
            cell.alignment = center
            cell.border = thin_border

    # 信息行
    duration_str = format_duration(per_student_seconds)
    info_row = len(rows) + 3
    ws.cell(row=info_row, column=1, value="⏱ 每人答辩时间：").font = Font(name="微软雅黑", bold=True, size=10.5)
    ws.cell(row=info_row, column=2, value=duration_str).font = Font(name="微软雅黑", size=10.5)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
