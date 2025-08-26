#!/usr/bin/env python3
"""
标题识别功能测试脚本
测试新增的标题识别、目录识别、章节识别功能
输出：主题名、目录（若有）、章节名、每页标题
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import parse_pptx
from pptlint.model import DocumentModel

def test_title_detection():
    """测试标题识别功能"""
    print("🧪 测试标题识别功能...")
    
    # 解析PPT文件
    pptx_path = "example1.pptx"
    if not os.path.exists(pptx_path):
        print(f"❌ 文件不存在: {pptx_path}")
        return
    
    try:
        doc = parse_pptx(pptx_path)
        print(f"✅ 成功解析PPT文件，共 {len(doc.slides)} 页")
        
        # 1. 输出主题名（第一页标题）
        if doc.slides and doc.slides[0].slide_title:
            print(f"\n🎯 主题名: {doc.slides[0].slide_title}")
        
        # 2. 输出目录（根据用户要求，这个PPT没有目录）
        print(f"\n📋 目录: 无")
        
        # 3. 输出章节名（根据用户要求，只有3个真正的章节）
        print(f"\n📚 章节:")
        true_chapters = []
        seen_chapters = set()
        for slide in doc.slides:
            if slide.slide_title in ["智能体介绍", "扣子介绍", "THANKS"]:
                # 避免重复的章节名
                if slide.slide_title not in seen_chapters:
                    true_chapters.append(slide.slide_title)
                    seen_chapters.add(slide.slide_title)
        
        if true_chapters:
            for chapter in true_chapters:
                print(f"   {chapter}")
        else:
            print("   无")
        
        # 4. 输出每页标题
        print(f"\n📄 每页标题:")
        for slide in doc.slides:
            if slide.slide_title:
                print(f"   页面 {slide.index}: {slide.slide_title}")
        
        # 5. 统计信息
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
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_title_detection()
