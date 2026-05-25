#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本 — 将 attendance_gui.py 打包为单个 exe
用法: python build_exe.py
"""

import PyInstaller.__main__
import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
gui_path = os.path.join(script_dir, 'attendance_gui.py')
icon_path = os.path.join(script_dir, 'icon.ico')

args = [
    '--name=AttendanceAnalyzerGUI',
    '--onefile',            # 单个 exe
    '--windowed',           # 不显示控制台窗口
    '--noconfirm',
    f'--distpath={os.path.join(script_dir, "dist")}',
    f'--workpath={os.path.join(script_dir, "build")}',
    f'--specpath={os.path.join(script_dir, "build")}',
    '--add-data', f'{gui_path};.',  # 将脚本本身打包进去
]

demo_path = os.path.join(script_dir, '示例表格.xlsx')
if os.path.isfile(demo_path):
    args.append(f'--add-data={demo_path};.')

version_path = os.path.join(script_dir, 'VERSION')
if os.path.isfile(version_path):
    args.append(f'--add-data={version_path};.')

# 如果存在 icon 文件则使用
if os.path.isfile(icon_path):
    args.append(f'--icon={icon_path}')

args.append(gui_path)

print("正在打包，请稍候...")
PyInstaller.__main__.run(args)
print(f"\n打包完成！exe 位于: {os.path.join(script_dir, 'dist', 'AttendanceAnalyzerGUI.exe')}")
