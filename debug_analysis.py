#!/usr/bin/env python3
"""
调试脚本：查看每个页面的详细分析结果
分析为什么某些页面没有被正确识别
"""

import sys
import os
import re
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import parse_pptx, _analyze_slide_content, _is_title_page, _is_toc_page, _is_chapter_page

def debug_analysis():
    """调试分析过程"""
    print("🔍 调试页面分析过程...")
    
    pptx_path = "example1.pptx"
    if not os.path.exists(pptx_path):
        print(f"❌ 文件不存在: {pptx_path}")
        return
    
    try:
        doc = parse_pptx(pptx_path)
        print(f"✅ 成功解析PPT文件，共 {len(doc.slides)} 页")
        
        # 第一遍扫描：收集目录内容
        toc_content = []
        for slide in doc.slides:
            for shape in slide.shapes:
                if shape.text_runs:
                    for text_run in shape.text_runs:
                        text = text_run.text.strip()
                        if text and re.match(r'^\d+[\.\s|｜]', text):
                            toc_content.append(text)
        
        print(f"\n📋 收集到的目录内容: {toc_content}")
        
        # 分析每个页面
        for i, slide in enumerate(doc.slides):
            print(f"\n{'='*50}")
            print(f"📄 页面 {i}: {slide.slide_title}")
            print(f"   实际类型: {slide.slide_type}")
            
            # 分析页面内容特征
            analysis = _analyze_slide_content(slide.shapes)
            print(f"\n🔍 内容分析:")
            print(f"   总文本长度: {analysis['total_text_length']}")
            print(f"   文本块数量: {analysis['text_blocks']}")
            print(f"   字体大小: {analysis['font_sizes']}")
            print(f"   有编号项目: {analysis['has_numbered_items']}")
            print(f"   编号模式: {analysis['numbered_patterns']}")
            print(f"   左上角文本: {[t['text'] for t in analysis['top_left_texts']]}")
            print(f"   大字体文本: {[t['text'] for t in analysis['large_font_texts']]}")
            
            # 测试各种判断函数
            print(f"\n🧪 判断结果:")
            is_title = _is_title_page(i, slide.shapes, analysis)
            is_toc = _is_toc_page(slide.shapes, analysis)
            is_chapter = _is_chapter_page(slide.shapes, analysis, toc_content)
            
            print(f"   标题页判断: {is_title}")
            print(f"   目录页判断: {is_toc}")
            print(f"   章节页判断: {is_chapter}")
            
            # 显示前几个文本块
            print(f"\n📝 前3个文本块:")
            text_count = 0
            for shape in slide.shapes:
                for text_run in shape.text_runs:
                    if text_count < 3 and text_run.text.strip():
                        print(f"     {text_run.text.strip()} (字体: {text_run.font_size_pt}pt)")
                        text_count += 1
                if text_count >= 3:
                    break
            
    except Exception as e:
        print(f"❌ 调试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_analysis()
