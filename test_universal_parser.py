#!/usr/bin/env python3
"""
测试通用解析器的通用性
验证新的解析器是否能处理不同类型的PPT文件
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

from pptlint.parser import parse_pptx

def test_universal_parser():
    """测试通用解析器"""
    print("🧪 测试通用解析器...")
    
    # 测试当前PPT文件
    pptx_path = "example1.pptx"
    if os.path.exists(pptx_path):
        print(f"\n📄 测试文件1: {pptx_path}")
        try:
            doc = parse_pptx(pptx_path)
            print(f"✅ 成功解析，共 {len(doc.slides)} 页")
            
            # 分析结构
            title_pages = [s for s in doc.slides if s.slide_type == "title"]
            toc_pages = [s for s in doc.slides if s.slide_type == "toc"]
            chapter_pages = [s for s in doc.slides if s.slide_type == "chapter"]
            section_pages = [s for s in doc.slides if s.slide_type == "section"]
            content_pages = [s for s in doc.slides if s.slide_type == "content"]
            
            print(f"   标题页: {len(title_pages)} 页")
            print(f"   目录页: {len(toc_pages)} 页")
            print(f"   章节页: {len(chapter_pages)} 页")
            print(f"   章节页(sec): {len(section_pages)} 页")
            print(f"   内容页: {len(content_pages)} 页")
            
            # 显示前几页的标题和类型
            print(f"\n📋 前5页结构:")
            for i in range(min(5, len(doc.slides))):
                slide = doc.slides[i]
                print(f"   页面 {i}: {slide.slide_title} ({slide.slide_type})")
                
        except Exception as e:
            print(f"❌ 解析失败: {e}")
    
    # 检查是否有其他PPT文件可以测试
    pptx_files = [f for f in os.listdir('.') if f.endswith('.pptx') and f != 'example1.pptx']
    
    if pptx_files:
        print(f"\n🔍 发现其他PPT文件: {pptx_files}")
        for pptx_file in pptx_files[:2]:  # 只测试前2个
            print(f"\n📄 测试文件: {pptx_file}")
            try:
                doc = parse_pptx(pptx_file)
                print(f"✅ 成功解析，共 {len(doc.slides)} 页")
                
                # 分析结构
                title_pages = [s for s in doc.slides if s.slide_type == "title"]
                toc_pages = [s for s in doc.slides if s.slide_type == "toc"]
                chapter_pages = [s for s in doc.slides if s.slide_type == "chapter"]
                section_pages = [s for s in doc.slides if s.slide_type == "section"]
                content_pages = [s for s in doc.slides if s.slide_type == "content"]
                
                print(f"   标题页: {len(title_pages)} 页")
                print(f"   目录页: {len(toc_pages)} 页")
                print(f"   章节页: {len(chapter_pages)} 页")
                print(f"   章节页(sec): {len(section_pages)} 页")
                print(f"   内容页: {len(content_pages)} 页")
                
                # 显示前3页的标题和类型
                print(f"   前3页结构:")
                for i in range(min(3, len(doc.slides))):
                    slide = doc.slides[i]
                    print(f"     页面 {i}: {slide.slide_title} ({slide.slide_type})")
                    
            except Exception as e:
                print(f"❌ 解析失败: {e}")
    else:
        print(f"\n📝 没有发现其他PPT文件用于测试")
    
    print(f"\n🎯 通用解析器测试完成！")

if __name__ == "__main__":
    test_universal_parser()
