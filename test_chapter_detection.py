#!/usr/bin/env python3
"""
章节识别测试脚本
专门测试改进后的章节识别功能
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import parse_pptx
from pptlint.model import DocumentModel

def test_chapter_detection():
    """测试章节识别功能"""
    print("🧪 测试章节识别功能...")
    
    # 解析PPT文件
    pptx_path = "example1.pptx"
    if not os.path.exists(pptx_path):
        print(f"❌ 文件不存在: {pptx_path}")
        return
    
    try:
        doc = parse_pptx(pptx_path)
        print(f"✅ 成功解析PPT文件，共 {len(doc.slides)} 页")
        
        # 分析每页的识别结果
        print(f"\n📋 详细识别结果:")
        for slide in doc.slides:
            print(f"\n📄 页面 {slide.index}:")
            print(f"   类型: {slide.slide_type}")
            print(f"   标题: {slide.slide_title}")
            
            # 显示形状信息
            if slide.shapes:
                print(f"   形状数量: {len(slide.shapes)}")
                for i, shape in enumerate(slide.shapes[:3]):  # 只显示前3个形状
                    if shape.text_runs:
                        text_content = " | ".join([tr.text for tr in shape.text_runs[:2]])  # 只显示前2个文本运行
                        print(f"     形状 {shape.id}: {text_content[:50]}...")
                        if shape.is_title:
                            print(f"       -> 识别为标题 (级别: {shape.title_level})")
                        if shape.is_toc:
                            print(f"       -> 识别为目录")
        
        # 统计各类型页面
        title_pages = [s for s in doc.slides if s.slide_type == "title"]
        toc_pages = [s for s in doc.slides if s.slide_type == "toc"]
        chapter_pages = [s for s in doc.slides if s.slide_type == "chapter"]
        section_pages = [s for s in doc.slides if s.slide_type == "section"]
        content_pages = [s for s in doc.slides if s.slide_type == "content"]
        
        print(f"\n📊 页面类型统计:")
        print(f"   标题页: {len(title_pages)} 页")
        print(f"   目录页: {len(toc_pages)} 页")
        print(f"   章节页: {len(chapter_pages)} 页")
        print(f"   章节页(sec): {len(section_pages)} 页")
        print(f"   内容页: {len(content_pages)} 页")
        print(f"   总计: {len(doc.slides)} 页")
        
        # 显示章节页面
        if chapter_pages:
            print(f"\n🎯 识别为章节的页面:")
            for slide in chapter_pages:
                print(f"   页面 {slide.index}: {slide.slide_title}")
        
        # 显示目录页面
        if toc_pages:
            print(f"\n📋 识别为目录的页面:")
            for slide in toc_pages:
                print(f"   页面 {slide.index}: {slide.slide_title}")
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_chapter_detection()
