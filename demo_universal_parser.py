#!/usr/bin/env python3
"""
演示通用解析器的功能
展示新解析器如何通用地识别不同类型的PPT结构
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import parse_pptx

def demo_universal_parser():
    """演示通用解析器功能"""
    print("🚀 通用解析器功能演示")
    print("=" * 50)
    
    # 测试文件列表
    test_files = ["example1.pptx", "example2.pptx"]
    
    for pptx_file in test_files:
        if not os.path.exists(pptx_file):
            continue
            
        print(f"\n📄 分析文件: {pptx_file}")
        print("-" * 40)
        
        try:
            doc = parse_pptx(pptx_file)
            print(f"✅ 成功解析，共 {len(doc.slides)} 页")
            
            # 1. 主题识别
            if doc.slides and doc.slides[0].slide_title:
                print(f"\n🎯 主题识别:")
                print(f"   主题名: {doc.slides[0].slide_title}")
                print(f"   页面类型: {doc.slides[0].slide_type}")
            
            # 2. 目录识别
            toc_pages = [s for s in doc.slides if s.slide_type == "toc"]
            if toc_pages:
                print(f"\n📋 目录识别:")
                for slide in toc_pages:
                    print(f"   页面 {slide.index}: {slide.slide_title}")
            else:
                print(f"\n📋 目录识别: 无目录页")
            
            # 3. 章节识别
            chapter_pages = [s for s in doc.slides if s.slide_type == "chapter"]
            if chapter_pages:
                print(f"\n📚 章节识别:")
                for slide in chapter_pages:
                    print(f"   页面 {slide.index}: {slide.slide_title}")
            else:
                print(f"\n📚 章节识别: 无明确章节页")
            
            # 4. 每页标题识别
            print(f"\n📄 每页标题识别:")
            for i, slide in enumerate(doc.slides):
                page_type = slide.slide_type
                title = slide.slide_title or "无标题"
                print(f"   页面 {i}: {title} ({page_type})")
            
            # 5. 统计信息
            title_pages = [s for s in doc.slides if s.slide_type == "title"]
            section_pages = [s for s in doc.slides if s.slide_type == "section"]
            content_pages = [s for s in doc.slides if s.slide_type == "content"]
            
            print(f"\n📊 结构统计:")
            print(f"   标题页: {len(title_pages)} 页")
            print(f"   目录页: {len(toc_pages)} 页")
            print(f"   章节页: {len(chapter_pages)} 页")
            print(f"   章节页(sec): {len(section_pages)} 页")
            print(f"   内容页: {len(content_pages)} 页")
            print(f"   总计: {len(doc.slides)} 页")
            
            # 6. 特殊功能演示
            print(f"\n🔍 特殊功能演示:")
            
            # 标题占位符检测
            title_placeholders = []
            for slide in doc.slides:
                for shape in slide.shapes:
                    if shape.is_title and shape.title_level == 1:
                        title_placeholders.append(f"页面{slide.index}: {shape.id}")
            
            if title_placeholders:
                print(f"   标题占位符: {', '.join(title_placeholders)}")
            else:
                print(f"   标题占位符: 无")
            
            # 多级标题检测
            multi_level_titles = {}
            for slide in doc.slides:
                for shape in slide.shapes:
                    if shape.is_title and shape.title_level > 0:
                        level = shape.title_level
                        if level not in multi_level_titles:
                            multi_level_titles[level] = []
                        multi_level_titles[level].append(f"页面{slide.index}: {shape.id}")
            
            if multi_level_titles:
                print(f"   多级标题:")
                for level in sorted(multi_level_titles.keys()):
                    print(f"     H{level}: {', '.join(multi_level_titles[level])}")
            else:
                print(f"   多级标题: 无")
                
        except Exception as e:
            print(f"❌ 解析失败: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n🎯 演示完成！")
    print(f"\n💡 新解析器的优势:")
    print(f"   1. 通用性: 不依赖特定PPT内容，可处理任何PPT文件")
    print(f"   2. 智能识别: 基于占位符、字体特征、位置等自动识别标题")
    print(f"   3. 多级支持: 支持H1、H2、H3等多级标题识别")
    print(f"   4. 结构分析: 自动识别标题页、目录页、章节页、内容页")
    print(f"   5. 借鉴成熟方案: 基于 pptx2md 的成熟实现")

if __name__ == "__main__":
    demo_universal_parser()
