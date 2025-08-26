#!/usr/bin/env python3
"""
简单测试脚本：直接测试判断函数
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import _is_toc_page, _is_chapter_page

def test_simple():
    """简单测试"""
    print("🧪 简单测试判断函数...")
    
    # 模拟页面7的数据
    print("\n📄 测试页面7（思考）:")
    analysis_7 = {
        'total_text_length': 461,
        'text_blocks': 51,
        'font_sizes': [14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0, 14.0],
        'has_numbered_items': True,
        'numbered_patterns': ['1.', '2.', '3.', '4.', '5.', '6.'],
        'top_left_texts': [],
        'large_font_texts': []
    }
    
    print(f"   有编号项目: {analysis_7['has_numbered_items']}")
    print(f"   编号项目数量: {len(analysis_7['numbered_patterns'])}")
    print(f"   文本块数量: {analysis_7['text_blocks']}")
    print(f"   总文本长度: {analysis_7['total_text_length']}")
    
    is_toc_7 = _is_toc_page([], analysis_7)
    print(f"   目录页判断: {is_toc_7}")
    
    # 模拟页面8的数据
    print("\n📄 测试页面8（扣子介绍）:")
    analysis_8 = {
        'total_text_length': 6,
        'text_blocks': 2,
        'font_sizes': [],
        'has_numbered_items': False,
        'numbered_patterns': [],
        'top_left_texts': [],
        'large_font_texts': []
    }
    
    print(f"   总文本长度: {analysis_8['total_text_length']}")
    print(f"   文本块数量: {analysis_8['text_blocks']}")
    print(f"   字体大小: {analysis_8['font_sizes']}")
    
    is_chapter_8 = _is_chapter_page([], analysis_8, [])
    print(f"   章节页判断: {is_chapter_8}")
    
    # 模拟页面12的数据
    print("\n📄 测试页面12（THANKS）:")
    analysis_12 = {
        'total_text_length': 15,
        'text_blocks': 2,
        'font_sizes': [60.0, 14.0],
        'has_numbered_items': False,
        'numbered_patterns': [],
        'top_left_texts': [],
        'large_font_texts': ['THANKS']
    }
    
    print(f"   总文本长度: {analysis_12['total_text_length']}")
    print(f"   文本块数量: {analysis_12['text_blocks']}")
    print(f"   字体大小: {analysis_12['font_sizes']}")
    
    is_chapter_12 = _is_chapter_page([], analysis_12, [])
    print(f"   章节页判断: {is_chapter_12}")

if __name__ == "__main__":
    test_simple()
