#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
雨课堂考勤数据处理
读取雨课堂导出的Excel文件（汇总页+子表），
生成每名学生在每次课的考勤状态：上课/旷课/病假/事假
"""

import sys
import os
from collections import OrderedDict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

INPUT_FILE = '大数据技术-2026春-计算机2305-计算机230---汇总-数据表-20260523144705_8716344.xlsx'
OUTPUT_FILE = '学生考勤明细表.xlsx'


def col_idx(col):
    i = 0
    for ch in col:
        i = i * 26 + (ord(ch) - 64)
    return i - 1


def parse_summary(wb):
    ws = wb[wb.sheetnames[0]]

    # Row 2: sub-headers
    row2 = {}
    for cell in ws[2]:
        if cell.value:
            row2[cell.column_letter] = str(cell.value)

    # Row 1: session headers (merged cells—only first cell has value)
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
                rec['scores'][sk] = int(float(raw)) if raw else None
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
                    pass  # skip MergedCell
            sid = cells.get('A', '')
            remark = cells.get('G', '')
            if sid and remark in ('病假', '事假'):
                leave_data[(session_key, sid)] = remark

    return leave_data


def main():
    dir_path = os.path.dirname(os.path.abspath(__file__))
    input_name = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    input_path = os.path.join(dir_path, input_name)
    output_path = os.path.join(dir_path, OUTPUT_FILE)

    import time
    output_path = os.path.join(dir_path, f'考勤明细_{time.strftime("%Y%m%d_%H%M%S")}.xlsx')

    if not os.path.exists(input_path):
        print(f"❌ 未找到文件: {input_name}")
        return

    wb = load_workbook(input_path, data_only=True)
    session_keys, session_sign_map, session_score_map, students = parse_summary(wb)
    leave_data = parse_sub_sheets(wb, session_keys)
    wb.close()

    # ── Sheet 1: 考勤明细 ──
    session_count = len(session_sign_map)
    headers = ['学号', '姓名'] + list(session_sign_map.keys()) + ['无故旷课率', '总旷课率']
    out_rows = []
    present_set = {'扫二维码', '“正在上课”提示', '教师添加', '课堂暗号'}

    for s in students:
        row = [s['id'], s['name']]
        absent_cnt = 0
        excused_cnt = 0  # 病假 + 事假
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

    # ── Sheet 1: write with colors ──
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    color_map = {'上课': green_fill, '病假': yellow_fill, '事假': yellow_fill, '旷课': red_fill}

    owb = Workbook()
    ows = owb.active
    ows.title = '考勤明细'
    ows.append(headers)
    for row in out_rows:
        ows.append(row)
    # Apply colors to data cells (skip 学号/姓名 and the stat columns)
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

    # Highlight top 10% students by total score
    sorted_totals = sorted(totals, reverse=True)
    n_top = max(1, round(len(sorted_totals) * 0.1))
    threshold = sorted_totals[n_top - 1] if n_top <= len(sorted_totals) else sorted_totals[-1]
    gold_fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')
    for row in ows2.iter_rows(min_row=2, max_row=ows2.max_row):
        if row[-1].value is not None and row[-1].value >= threshold:
            for cell in row:
                cell.fill = gold_fill

    owb.save(output_path)
    print(f"  -> 已生成: {output_path}")

    # Summary
    print(f"   共 {len(students)} 名学生, {len(session_sign_map)} 次课\n")
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
        print(f"  {sk:20s}  上课{attend:>2}人  旷课{absent:>2}人  病假{sick:>2}人  事假{personal:>2}人")

    total_absent = sum(1 for s in students for sk in session_sign_map
                       if s['attendance'].get(sk) == '未上课'
                       and (sk, s['id']) not in leave_data)
    total_sick = sum(1 for v in leave_data.values() if v == '病假')
    total_personal = sum(1 for v in leave_data.values() if v == '事假')
    print(f"\n  旷课总人次: {total_absent}")
    print(f"  病假总人次: {total_sick}")
    print(f"  事假总人次: {total_personal}")


if __name__ == '__main__':
    main()
