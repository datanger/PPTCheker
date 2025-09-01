#!/usr/bin/env python3
"""
PPT审查工具 - GUI启动脚本
"""
import sys
import os

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入并运行GUI
from app.pptlint.simple_gui import main

if __name__ == "__main__":
    main()
