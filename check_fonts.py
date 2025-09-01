#!/usr/bin/env python3
"""
字体检测脚本 - 检查Ubuntu系统上的中文字体
"""
import tkinter as tk
from tkinter import font
import subprocess
import os

def check_system_fonts():
    """检查系统字体"""
    print("=== 系统字体检测 ===")
    
    # 检查fc-list命令
    try:
        result = subprocess.run(['fc-list', ':lang=zh'], capture_output=True, text=True)
        if result.returncode == 0:
            print("系统中文字体列表:")
            for line in result.stdout.split('\n')[:10]:  # 只显示前10个
                if line.strip():
                    print(f"  {line.strip()}")
        else:
            print("无法获取字体列表")
    except FileNotFoundError:
        print("fc-list命令不可用")
    
    print("\n=== Tkinter字体检测 ===")
    
    # 创建临时窗口检测字体
    root = tk.Tk()
    root.withdraw()  # 隐藏窗口
    
    # 测试字体列表
    test_fonts = [
        'WenQuanYi Micro Hei',
        'WenQuanYi Zen Hei', 
        'Noto Sans CJK SC',
        'Noto Sans CJK TC',
        'Source Han Sans CN',
        'Ubuntu',
        'DejaVu Sans',
        'Liberation Sans',
        'Arial',
        'TkDefaultFont',
        'TkHeadingFont',
        'TkFixedFont'
    ]
    
    print("Tkinter可用字体:")
    for font_name in test_fonts:
        try:
            # 测试字体
            test_font = font.Font(family=font_name, size=9)
            test_result = test_font.measure('测试')
            if test_result > 0:
                print(f"  ✓ {font_name} - 可用")
            else:
                print(f"  ✗ {font_name} - 不可用")
        except Exception as e:
            print(f"  ✗ {font_name} - 错误: {e}")
    
    # 获取系统默认字体
    try:
        default_font = font.nametofont('TkDefaultFont')
        print(f"\n系统默认字体: {default_font.actual()}")
    except:
        print("\n无法获取系统默认字体")
    
    root.destroy()

if __name__ == "__main__":
    check_system_fonts()
