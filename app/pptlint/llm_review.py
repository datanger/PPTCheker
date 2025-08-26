"""
基于大模型的智能PPT审查模块

设计理念：
- 将PPT内容转换为结构化文本，让LLM进行语义分析
- 支持多种审查维度：格式规范、内容逻辑、术语一致性、表达流畅性
- 提供具体的修复建议和改进方案
"""
import json
from typing import List, Dict, Any, Optional
from .model import DocumentModel, Issue, TextRun
from .llm import LLMClient
from .config import ToolConfig


class LLMReviewer:
    """基于LLM的智能审查器"""
    
    def __init__(self, llm: LLMClient, config: ToolConfig):
        self.llm = llm
        self.config = config
        
    def is_enabled(self) -> bool:
        """检查LLM是否可用"""
        return self.llm.is_enabled()
    
    def extract_slide_content(self, doc: DocumentModel) -> List[Dict[str, Any]]:
        """提取幻灯片内容，转换为LLM可理解的格式"""
        slides_content = []
        
        for slide in doc.slides:
            slide_data = {
                "slide_index": slide.index,
                "slide_title": slide.slide_title,
                "slide_type": slide.slide_type,
                "chapter_info": slide.chapter_info,
                "text_blocks": [],
                "titles": [],
                "fonts": set(),
                "colors": set(),
                "raw_text": ""
            }
            
            for shape in slide.shapes:
                for text_run in shape.text_runs:
                    if text_run.text.strip():
                        block = {
                            "text": text_run.text,
                            "font": text_run.font_name,
                            "size": text_run.font_size_pt,
                            "language": text_run.language_tag,
                            "shape_id": shape.id,
                            "is_title": shape.is_title,
                            "title_level": shape.title_level,
                            "is_bold": text_run.is_bold,
                            "is_italic": text_run.is_italic,
                            "is_underline": text_run.is_underline
                        }
                        slide_data["text_blocks"].append(block)
                        slide_data["raw_text"] += text_run.text + " "
                        
                        # 收集标题信息
                        if shape.is_title and shape.title_level:
                            slide_data["titles"].append({
                                "text": text_run.text,
                                "level": shape.title_level,
                                "font": text_run.font_name,
                                "size": text_run.font_size_pt,
                                "is_bold": text_run.is_bold
                            })
                        
                        if text_run.font_name:
                            slide_data["fonts"].add(text_run.font_name)
                        if text_run.font_size_pt:
                            slide_data["colors"].add(text_run.font_size_pt)
            
            # 将set转换为list，确保JSON序列化
            slide_data["fonts"] = list(slide_data["fonts"])
            slide_data["colors"] = list(slide_data["colors"])
            
            slides_content.append(slide_data)
            
        return slides_content
    
    def review_format_standards(self, slides_content: List[Dict]) -> List[Issue]:
        """审查格式标准：字体、字号、颜色等"""
        if not self.is_enabled():
            return []
            
        prompt = f"""
            你是一个专业的PPT格式审查专家。请分析以下PPT内容，检查格式规范问题：

            审查标准：
            - 日文字体：应使用 {self.config.jp_font_name}
            - 最小字号：{self.config.min_font_size_pt}pt
            - 单页颜色数：不超过{self.config.color_count_threshold}种

            PPT内容：
            {json.dumps(slides_content, ensure_ascii=False, indent=2)}

            **重要**：请为每个问题提供页面级别的对象引用，格式如下：
            - 如果问题影响整个页面：使用 "page_[页码]"
            - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_FormatRule",
                        "severity": "warning|info",
                        "slide_index": 0,
                        "object_ref": "page_0",
                        "message": "问题描述",
                        "suggestion": "具体建议",
                        "can_autofix": true|false
                    }}
                ]
            }}

            只返回JSON，不要其他内容。
            """
        
        try:
            response = self.llm.complete(prompt, max_tokens=1024)
            if response:
                # 尝试解析JSON响应
                result = json.loads(response.strip())
                issues = []
                
                for item in result.get("issues", []):
                    issue = Issue(
                        file="",  # 会在workflow中设置
                        slide_index=item.get("slide_index", 0),
                        object_ref=item.get("object_ref", "page"),
                        rule_id=item.get("rule_id", "LLM_FormatRule"),
                        severity=item.get("severity", "info"),
                        message=item.get("message", ""),
                        suggestion=item.get("suggestion", ""),
                        can_autofix=item.get("can_autofix", False)
                    )
                    issues.append(issue)
                
                return issues
        except Exception as e:
            print(f"LLM格式审查失败: {e}")
            
        return []
    
    def review_content_logic(self, slides_content: List[Dict]) -> List[Issue]:
        """审查内容逻辑：连贯性、术语一致性、表达流畅性"""
        if not self.is_enabled():
            return []
            
        prompt = f"""
            你是一个专业的PPT内容审查专家。请分析以下PPT内容，检查内容逻辑问题：

            审查维度：
            1. 逻辑连贯性：各页面之间的逻辑过渡是否自然
            2. 术语一致性：相同概念是否使用统一术语
            3. 表达流畅性：语言表达是否清晰准确
            4. 内容完整性：是否遗漏重要信息

            PPT内容：
            {json.dumps(slides_content, ensure_ascii=False, indent=2)}

            **重要**：请为每个问题提供页面级别的对象引用，格式如下：
            - 如果问题影响整个页面：使用 "page_[页码]"
            - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_ContentRule",
                        "severity": "warning|info",
                        "slide_index": 0,
                        "object_ref": "page_0",
                        "message": "问题描述",
                        "suggestion": "具体建议",
                        "can_autofix": false
                    }}
                ]
            }}

            只返回JSON，不要其他内容。
            """
        
        try:
            response = self.llm.complete(prompt, max_tokens=1024)
            if response:
                result = json.loads(response.strip())
                issues = []
                
                for item in result.get("issues", []):
                    issue = Issue(
                        file="",
                        slide_index=item.get("slide_index", 0),
                        object_ref=item.get("object_ref", "page"),
                        rule_id=item.get("rule_id", "LLM_ContentRule"),
                        severity=item.get("severity", "info"),
                        message=item.get("message", ""),
                        suggestion=item.get("suggestion", ""),
                        can_autofix=item.get("can_autofix", False)
                    )
                    issues.append(issue)
                
                return issues
        except Exception as e:
            print(f"LLM内容审查失败: {e}")
            
        return []
    
    def review_acronyms(self, slides_content: List[Dict]) -> List[Issue]:
        """智能审查缩略语：基于LLM理解上下文，只标记真正需要解释的缩略语"""
        if not self.is_enabled():
            print("    LLM未启用，跳过缩略语审查")
            return []
            
        print(f"    🧠 开始缩略语审查，分析 {len(slides_content)} 个页面...")
            
        prompt = f"""
            你是一个专业的PPT内容审查专家，专门负责缩略语使用审查。

            审查原则：
            1. **常见缩略语不需要解释**：如API、URL、HTTP、HTML、CSS、JS、SQL、GUI、CLI、IDE、SDK、CPU、GPU、RAM、USB、WiFi、GPS、TV、DVD、CD、MP3、MP4、PDF、PPT、AI、ML、DL、VR、AR、IoT、CEO、CTO、CFO、HR、IT、PR、QA、UI、UX、PM、USA、UK、EU、UN、WHO、NASA、FBI、CIA、THANKS、OK、FAQ、ASAP、FYI、IMO、BTW、LOL、OMG等
            2. **专业术语缩略语需要解释**：如LLM（Large Language Model）、MCP（Model Context Protocol）、UFO（User-Friendly Operating system）等
            3. **判断标准**：基于目标读者群体（假设是IT行业专业人士）的知识水平来判断

            PPT内容：
            {json.dumps(slides_content, ensure_ascii=False, indent=2)}

            请分析每个缩略语，判断是否需要解释。只标记那些：
            - 目标读者可能不理解的
            - 首次出现且缺乏解释的
            - 专业性强或行业特定的

            **重要**：请为每个问题提供精确的对象引用，格式如下：
            - 如果缩略语在特定文本块中：使用 "text_block_[页码]_[块索引]"
            - 如果缩略语在页面标题中：使用 "title_[页码]"
            - 如果缩略语在页面级别且无法精确定位：使用 "page_[页码]"

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_AcronymRule",
                        "severity": "info",
                        "slide_index": 0,
                        "object_ref": "text_block_0_1",
                        "message": "专业缩略语 [缩略语名称] 首次出现未发现解释",
                        "suggestion": "建议在首次出现后添加解释：[缩略语名称] (全称)",
                        "can_autofix": false
                    }}
                ]
            }}

            只返回JSON，不要其他内容。
            """
        
        try:
            print(f"    📤 发送LLM请求...")
            response = self.llm.complete(prompt, max_tokens=1024)
            print(f"    📥 收到LLM响应: {response[:100] if response else 'None'}...")
            
            if response:
                result = json.loads(response.strip())
                issues = []
                
                for item in result.get("issues", []):
                    issue = Issue(
                        file="",
                        slide_index=item.get("slide_index", 0),
                        object_ref=item.get("object_ref", "page"),
                        rule_id=item.get("rule_id", "LLM_AcronymRule"),
                        severity=item.get("severity", "info"),
                        message=item.get("message", ""),
                        suggestion=item.get("suggestion", ""),
                        can_autofix=item.get("can_autofix", False)
                    )
                    issues.append(issue)
                
                print(f"    ✅ 缩略语审查完成，发现 {len(issues)} 个问题")
                return issues
            else:
                print(f"    ⚠️ LLM响应为空")
                return []
        except Exception as e:
            print(f"    ❌ LLM缩略语审查失败: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def review_title_structure(self, slides_content: List[Dict]) -> List[Issue]:
        """审查标题结构：目录、章节、页面标题的层级一致性和逻辑连贯性"""
        if not self.is_enabled():
            return []
            
        print("    📋 审查标题结构...")
        
        prompt = f"""
            你是一个专业的PPT标题结构审查专家。请分析以下PPT内容，检查标题结构问题：

            审查维度：
            1. **目录识别**：识别目录页面，检查目录项的完整性和准确性
            2. **章节结构**：检查章节标题的层级关系（H1/H2/H3）是否合理
            3. **标题一致性**：检查标题的命名风格、格式是否统一
            4. **逻辑连贯性**：检查标题之间的逻辑关系和过渡是否自然
            5. **页面标题**：检查每页标题是否清晰、准确反映页面内容

            PPT内容：
            {json.dumps(slides_content, ensure_ascii=False, indent=2)}

            **重要**：请为每个问题提供精确的对象引用，格式如下：
            - 如果问题影响整个页面：使用 "page_[页码]"
            - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"
            - 如果问题在标题中：使用 "title_[页码]"

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_TitleStructureRule",
                        "severity": "warning|info",
                        "slide_index": 0,
                        "object_ref": "title_0",
                        "message": "问题描述",
                        "suggestion": "具体建议",
                        "can_autofix": false
                    }}
                ]
            }}

            只返回JSON，不要其他内容。
            """
        
        try:
            response = self.llm.complete(prompt, max_tokens=1024)
            if response:
                result = json.loads(response.strip())
                issues = []
                
                for item in result.get("issues", []):
                    issue = Issue(
                        file="",
                        slide_index=item.get("slide_index", 0),
                        object_ref=item.get("object_ref", "page"),
                        rule_id=item.get("rule_id", "LLM_TitleStructureRule"),
                        severity=item.get("severity", "info"),
                        message=item.get("message", ""),
                        suggestion=item.get("suggestion", ""),
                        can_autofix=item.get("can_autofix", False)
                    )
                    issues.append(issue)
                
                print(f"    ✅ 标题结构审查完成，发现 {len(issues)} 个问题")
                return issues
            else:
                print(f"    ⚠️ LLM响应为空")
                return []
        except Exception as e:
            print(f"    ❌ LLM标题结构审查失败: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def run_llm_review(self, doc: DocumentModel) -> List[Issue]:
        """运行完整的LLM审查流程"""
        if not self.is_enabled():
            return []
            
        print("🤖 启动LLM智能审查...")
        
        # 提取内容
        slides_content = self.extract_slide_content(doc)
        
        # 多维度审查
        all_issues = []
        
        # 1. 格式标准审查
        print("📝 审查格式标准...")
        format_issues = self.review_format_standards(slides_content)
        all_issues.extend(format_issues)
        
        # 2. 内容逻辑审查
        print("🧠 审查内容逻辑...")
        logic_issues = self.review_content_logic(slides_content)
        all_issues.extend(logic_issues)
        
        # 3. 缩略语审查
        print("🔤 审查缩略语...")
        acronym_issues = self.review_acronyms(slides_content)
        all_issues.extend(acronym_issues)
        
        # 4. 标题结构审查
        print("📋 审查标题结构...")
        title_structure_issues = self.review_title_structure(slides_content)
        all_issues.extend(title_structure_issues)
        
        print(f"✅ LLM审查完成，发现 {len(all_issues)} 个问题")
        return all_issues


def create_llm_reviewer(llm: LLMClient, config: ToolConfig) -> LLMReviewer:
    """创建LLM审查器实例"""
    return LLMReviewer(llm, config)
