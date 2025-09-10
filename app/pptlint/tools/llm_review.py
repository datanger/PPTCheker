"""
基于大模型的智能PPT审查模块

设计理念：
- 将PPT内容转换为结构化文本，让LLM进行语义分析
- 支持多种审查维度：格式规范、内容逻辑、术语一致性、表达流畅性
- 提供具体的修复建议和改进方案
"""
import json
from typing import List, Dict, Any, Optional
try:
    from ..model import DocumentModel, Issue, TextRun
    from ..llm import LLMClient
    from ..config import ToolConfig
except ImportError:
    # 兼容直接运行的情况
    import sys
    import os
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from model import DocumentModel, Issue, TextRun
    from llm import LLMClient
    from config import ToolConfig


class LLMReviewer:
    """基于LLM的智能审查器"""
    
    def __init__(self, llm: LLMClient, config: ToolConfig):
        self.llm = llm
        self.config = config
        self.stop_event = None  # 停止事件
        # 导入提示词管理器
        try:
            from ..prompt_manager import prompt_manager
            self.prompt_manager = prompt_manager
        except ImportError:
            self.prompt_manager = None
    
    def set_stop_event(self, stop_event):
        """设置停止事件"""
        self.stop_event = stop_event
    
    def _clean_json_response(self, response: str) -> str:
        """清理LLM响应中的markdown代码块标记和其他格式问题"""
        if not response or not response.strip():
            return ""
            
        cleaned_response = response.strip()
        
        # 移除markdown代码块标记
        if cleaned_response.startswith('```json'):
            cleaned_response = cleaned_response[7:]
        elif cleaned_response.startswith('```'):
            cleaned_response = cleaned_response[3:]
            
        if cleaned_response.endswith('```'):
            cleaned_response = cleaned_response[:-3]
            
        cleaned_response = cleaned_response.strip()
        
        # 移除可能的其他前缀
        prefixes_to_remove = [
            "JSON格式：",
            "JSON:",
            "json:",
            "返回结果：",
            "结果：",
            "Response:",
            "response:",
        ]
        
        for prefix in prefixes_to_remove:
            if cleaned_response.startswith(prefix):
                cleaned_response = cleaned_response[len(prefix):].strip()
                break
        
        # 查找JSON对象的开始和结束
        start_idx = cleaned_response.find('{')
        end_idx = cleaned_response.rfind('}')
        
        if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
            cleaned_response = cleaned_response[start_idx:end_idx + 1]
        
        return cleaned_response.strip()
    
    def _get_default_report_optimization_prompt(self, report_md: str, issues: List[Issue]) -> str:
        """获取默认报告优化提示词"""
        return f"""
你是一个专业的报告优化专家。请对以下PPT审查报告进行优化，主要目标是：

**优化要求：**
1. **删除重复内容**：如果同一个问题在多个页面重复出现，只保留最重要的1-2个实例
2. **精简无关紧要的提示**：删除过于琐碎或不重要的建议，只保留核心问题
3. **合并相似问题**：将相同类型的问题合并为一条，避免重复
4. **保持报告结构**：维持原有的Markdown格式和层次结构
5. **突出重要问题**：确保严重问题（serious级别）得到突出显示

**原始报告：**
```markdown
{report_md}
```

**问题统计：**
- 总问题数：{len(issues)}
- 规则问题：{len([i for i in issues if not i.rule_id.startswith('LLM_')])}
- LLM问题：{len([i for i in issues if i.rule_id.startswith('LLM_')])}

请返回优化后的报告，保持Markdown格式，确保：
- 删除重复和冗余内容
- 保留所有重要问题
- 维持清晰的层次结构
- 突出关键改进建议

只返回优化后的Markdown报告，不要其他内容。
"""
    
    def _get_default_format_prompt(self, pages: List[Dict[str, Any]]) -> str:
        """获取默认格式审查提示词"""
        return f"""
            你是一个专业的PPT格式审查专家。请分析以下PPT内容，检查格式规范问题：

            审查标准：
            - 日文字体：应使用 {self.config.jp_font_name}
            - 最小字号：{self.config.min_font_size_pt}pt
            - 单页颜色数：不超过{self.config.color_count_threshold}种

            PPT内容：
            {json.dumps(pages, ensure_ascii=False, indent=2)}

            **重要**：请为每个问题提供页面级别的对象引用，格式如下：
            - 如果问题影响整个页面：使用 "page_[页码]"
            - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_FormatRule",
                        "severity": "warning|info|serious",
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
    
    def _get_default_content_logic_prompt(self, parsing_data: Dict[str, Any]) -> str:
        """获取默认内容逻辑审查提示词"""
        return f"""
            你是一位非常挑剔和严谨的公司高层领导，正在审核下属提交的PPT汇报材料。你的标准极其严格，不容许任何逻辑漏洞、表达不清或结构混乱的问题。

            作为挑剔的领导，请从以下维度严格审查PPT内容：

            **1. 页内逻辑连贯性（极其重要）**
            - 每页内的标题、要点、图表是否逻辑清晰，层次分明
            - 页面内容是否围绕核心主题展开，避免无关信息
            - 要点之间是否有清晰的逻辑关系（并列、递进、因果等）
            - 是否存在逻辑跳跃、思维混乱的问题

            **2. 跨页逻辑连贯性（极其重要）**
            - 页面之间的过渡是否自然流畅，避免突兀的跳跃
            - 标题层级是否合理，章节结构是否清晰
            - 前后页面是否存在逻辑断层或重复冗余
            - 整体叙述线索是否清晰，听众能否跟上思路
            - 跨页的逻辑检查参考structure这个字段, 通过PPT的结构来判断跨页的逻辑是否连贯, 是否没有围绕核心主题展开

            **3. 标题与内容一致性（极其重要）**
            - 页面标题是否准确反映页面内容
            - 章节标题是否与内容要点匹配
            - 是否存在标题与内容不符的问题
            - 标题层级是否合理，避免混乱

            **4. 术语表达严谨性**
            - 专业术语使用是否一致，避免同一概念用不同词汇
            - 表达是否准确清晰，避免模糊不清的表述
            - 是否存在歧义或容易误解的表达
            - 特别需要检查语言表达是否符合该语种表达习惯，尤其是日语需要重点关注，若发现不符合表达习惯，则标记为问题

            **5. 内容结构完整性**
            - 是否遗漏关键信息或重要步骤
            - 各部分内容是否平衡，重点是否突出
            - 是否存在内容重复或冗余
            - **页面内容完整性**：有标题的页面是否包含相应的内容
            - **空内容页面检查**：是否存在只有标题但内容为空或过少的页面（如只有标题占位符，没有实际内容）

            **审查标准（极其严格）**：
            - 以挑剔领导的视角，找出任何可能影响汇报效果的问题
            - 重点关注逻辑连贯性，不容许任何跳跃或混乱
            - 对表达不清、结构混乱的问题零容忍
            - 对标题与内容不符的问题零容忍
            - **对空内容页面零容忍**：有标题但内容为空或过少的页面是严重问题

            **重要**：请为每个问题提供精确的对象引用，格式如下：
            - 如果问题影响整个页面：使用 "page_[页码]"
            - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"
            - 如果问题涉及标题：使用 "title_[页码]"

            PPT完整数据：
            {json.dumps(parsing_data, ensure_ascii=False, indent=2)}

            请以JSON格式返回审查结果，格式如下：
            {{
                "issues": [
                    {{
                        "rule_id": "LLM_ContentRule",
                        "severity": "warning|info|serious",
                        "slide_index": 1（注意：页码从1开始计数）,
                        "object_ref": "page_1（注意：页码从1开始计数）",
                        "message": "问题描述（要具体、明确、一针见血）",
                        "suggestion": "具体建议（要实用、可操作）",
                        "can_autofix": false
                    }}
                ]
            }}

            只返回JSON，不要其他内容。
            """
        
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
    
    def review_format_standards(self, parsing_data: Dict[str, Any]) -> List[Issue]:
        """审查格式标准：字体、字号、颜色等"""
        # 提取页面内容
        pages = parsing_data.get("contents", [])
        
        # 使用提示词管理器获取用户提示词
        if self.prompt_manager:
            user_prompt = self.prompt_manager.get_user_prompt_for_review(
                'format_standards',
                jp_font_name=self.config.jp_font_name,
                min_font_size_pt=self.config.min_font_size_pt,
                color_count_threshold=self.config.color_count_threshold
            )
            # 构建完整提示词：用户提示 + 输入提示 + 输出提示
            prompt = f"""{user_prompt}

PPT内容：
{json.dumps(pages, ensure_ascii=False, indent=2)}

**重要**：请为每个问题提供页面级别的对象引用，格式如下：
- 如果问题影响整个页面：使用 "page_[页码]"
- 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"

请以JSON格式返回审查结果，格式如下：
{{
    "issues": [
        {{
            "rule_id": "LLM_FormatRule",
            "severity": "warning|info|serious",
            "slide_index": 0,
            "object_ref": "page_0",
            "message": "问题描述",
            "suggestion": "具体建议",
            "can_autofix": true|false
        }}
    ]
}}

只返回JSON，不要其他内容。"""
        else:
            # 回退到默认提示词
            prompt = self._get_default_format_prompt(pages)
        
        try:
            response = self.llm.complete(prompt, max_tokens=self.config.llm_max_tokens, stop_event=self.stop_event)
            if response:
                # 尝试解析JSON响应
                cleaned_response = self._clean_json_response(response)
                try:
                    result = json.loads(cleaned_response)
                except json.JSONDecodeError as e:
                    print(f"    ❌ JSON解析失败: {e}")
                    print(f"    📄 清理后的响应: {cleaned_response}")
                    return []
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
    
    def review_content_logic(self, parsing_data: Dict[str, Any]) -> List[Issue]:
        """审查内容逻辑：连贯性、术语一致性、表达流畅性"""
        
        # 使用提示词管理器获取用户提示词
        if self.prompt_manager:
            user_prompt = self.prompt_manager.get_user_prompt_for_review('content_logic')
            # 构建完整提示词：用户提示 + 输入提示 + 输出提示
            prompt = f"""{user_prompt}

                **重要**：请为每个问题提供精确的对象引用，格式如下：
                - 如果问题影响整个页面：使用 "page_[页码]"
                - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"
                - 如果问题涉及标题：使用 "title_[页码]"

                PPT完整数据：
                {json.dumps(parsing_data, ensure_ascii=False, indent=2)}

                请以JSON格式返回审查结果，格式如下：
                {{
                    "issues": [
                        {{
                            "rule_id": "LLM_ContentRule",
                            "severity": "warning|info|serious",
                            "slide_index": 1（注意：页码从1开始计数）,
                            "object_ref": "page_1（注意：页码从1开始计数）",
                            "message": "问题描述（要具体、明确、一针见血）",
                            "suggestion": "具体建议（要实用、可操作）",
                            "can_autofix": false
                        }}
                    ]
                }}

                只返回JSON，不要其他内容。"""
        else:
            # 回退到默认提示词
            prompt = self._get_default_content_logic_prompt(parsing_data)
        
        try:
            print(f"    📤 发送LLM内容逻辑审查请求...")
            print(f"    🔑 使用模型: {self.llm.model}")
            print(f"    🌐 使用端点: {self.llm.endpoint}")
            print(f"    📝 提示词长度: {len(prompt)}")
            
            response = self.llm.complete(prompt, max_tokens=self.config.llm_max_tokens, stop_event=self.stop_event)
            print(f"    📥 收到LLM响应: {response[:200] if response else 'None'}...")
            print(f"    📏 响应长度: {len(response) if response else 0}")
            print(f"    🔍 响应类型: {type(response)}")
            print(f"    ✅ 响应非空: {bool(response)}")
            print(f"    ✅ 响应非空白: {bool(response and response.strip())}")
            
            if response and response.strip():
                try:
                    cleaned_response = self._clean_json_response(response)
                    try:
                        result = json.loads(cleaned_response)
                    except json.JSONDecodeError as e:
                        print(f"    ❌ JSON解析失败: {e}")
                        print(f"    📄 清理后的响应: {cleaned_response}")
                        return []
                    issues = []
                    
                    for item in result.get("issues", []):
                        # 处理页码：将LLM返回的页码（从1开始）转换为数组索引（从0开始）
                        slide_index = item.get("slide_index", 1)
                        array_index = max(0, slide_index - 1)  # 确保不会小于0
                        
                        issue = Issue(
                            file="",
                            slide_index=array_index,  # 使用数组索引（从0开始）
                            object_ref=item.get("object_ref", "page"),
                            rule_id=item.get("rule_id", "LLM_ContentRule"),
                            severity=item.get("severity", "info"),
                            message=item.get("message", ""),
                            suggestion=item.get("suggestion", ""),
                            can_autofix=item.get("can_autofix", False)
                        )
                        issues.append(issue)
                    print(f"    ✅ 内容逻辑审查完成，发现 {len(issues)} 个问题")
                    return issues
                except json.JSONDecodeError as e:
                    print(f"    ❌ JSON解析失败: {e}")
                    print(f"    📝 原始响应: {response[:500]}")
                except Exception as e:
                    print(f"    ❌ 处理响应失败: {e}")
            else:
                print(f"    ⚠️ LLM响应为空或无效")
                
        except Exception as e:
            print(f"    ❌ LLM内容审查失败: {e}")
            import traceback
            traceback.print_exc()
            
        return []
    
    def review_acronyms(self, parsing_data: Dict[str, Any]) -> List[Issue]:
        """智能审查缩略语：基于LLM理解上下文，只标记真正需要解释的缩略语"""
        # 提取页面内容
        pages = parsing_data.get("contents", [])
        print(f"    🧠 开始缩略语审查，分析 {len(pages)} 个页面...")
        
        # 使用提示词管理器获取用户提示词
        if self.prompt_manager:
            user_prompt = self.prompt_manager.get_user_prompt_for_review('acronyms')
            # 构建完整提示词：用户提示 + 输入提示 + 输出提示
            prompt = f"""{user_prompt}

                PPT内容：
                {json.dumps(pages, ensure_ascii=False, indent=2)}

                请分析每个缩略语，判断是否需要解释。只标记那些：
                - 目标读者可能不理解的
                - 首次出现且缺乏解释的
                - 专业性强或行业特定的
                - **重要**：如果同一页面内已经提供了该缩略语的解释，则不要标记
                - 如果某页之前已经解释过的缩略语，则不要标记
                - 针对某个缩略语不要重复标记，只针对第一次出现的位置进行标记

                主观评判标准：
                假设你是一个公司的高层领导在审查下面员工的PPT汇报材料，你不太懂专业领域术语，当在查看某页PPT时，看到某个缩略语不太懂其中的含义，但未在该页内找到解释，你认为需要解释，则标记为需要解释。

                **特别注意**：
                - 如果某页已经解释了某个缩略语（如"LLM：Large Language Model"），则不要标记该页的LLM
                - 优先标记那些没有解释的专业技术缩略语
                - 避免标记常见的逻辑词汇和基础术语

                **重要**：请仔细分析每个页面，准确识别缩略语所在的页面索引，页面索引从1开始计数。

                请以JSON格式返回审查结果，格式如下：
                {{
                    "issues": [
                        {{
                            "rule_id": "LLM_AcronymRule",
                            "severity": "info|warning|serious",
                            "slide_index": 1（注意替换成实际页码，从1开始计数）,
                            "object_ref": "page_1（注意替换成实际页码，从1开始计数）,
                            "message": "专业缩略语 [缩略语名称] 首次出现未发现解释",
                            "suggestion": "建议在首次出现后添加解释：[缩略语名称] (全称)",
                            "can_autofix": false
                        }}
                    ]
                }}

                只返回JSON，不要其他内容。"""
        else:
            # 回退到默认提示词
            prompt = self._get_default_acronyms_prompt(pages)
        
        try:
            print(f"    📤 发送LLM请求...")
            response = self.llm.complete(prompt, max_tokens=self.config.llm_max_tokens, stop_event=self.stop_event)
            print(f"    📥 收到LLM响应: {response[:100] if response else 'None'}...")
            
            if response:
                cleaned_response = self._clean_json_response(response)
                try:
                    result = json.loads(cleaned_response)
                except json.JSONDecodeError as e:
                    print(f"    ❌ JSON解析失败: {e}")
                    print(f"    📄 清理后的响应: {cleaned_response}")
                    return []
                issues = []
                
                for item in result.get("issues", []):
                    # 验证和纠正页面索引
                    slide_index = item.get("slide_index", 1)  # 默认从1开始
                    object_ref = item.get("object_ref", "page")
                    
                    # 将LLM返回的页码（从1开始）转换为数组索引（从0开始）
                    array_index = slide_index - 1
                    
                    # 如果LLM返回的页面索引超出范围，尝试自动纠正
                    if array_index < 0 or array_index >= len(pages):
                        print(f"    ⚠️ LLM返回的页面索引 {slide_index} 超出范围，尝试自动纠正...")
                        # 搜索整个PPT，找到包含相关缩略语的页面
                        corrected_index = self._find_acronym_page(pages, item.get("message", ""))
                        if corrected_index is not None:
                            array_index = corrected_index
                            slide_index = corrected_index + 1  # 转换回从1开始的页码
                            object_ref = f"page_{slide_index}"
                            print(f"    ✅ 自动纠正页面索引为: {slide_index} (数组索引: {array_index})")
                        else:
                            print(f"    ❌ 无法找到相关缩略语，跳过此问题")
                            continue
                    
                    # 使用转换后的数组索引创建Issue对象
                    issue = Issue(
                        file="",
                        slide_index=array_index,  # 使用数组索引（从0开始）
                        object_ref=object_ref,
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
    
    def _find_acronym_page(self, pages: List[Dict[str, Any]], message: str) -> Optional[int]:
        """搜索包含缩略语的页面索引"""
        try:
            # 从消息中提取缩略语名称
            import re
            acronym_match = re.search(r'\[([A-Z]+)\]', message)
            if not acronym_match:
                return None
            
            acronym = acronym_match.group(1)
            print(f"    🔍 搜索缩略语 '{acronym}' 所在的页面...")
            
            # 搜索每个页面
            for page_idx, page in enumerate(pages):
                # 检查页面标题
                page_title = page.get("页标题", "")
                if acronym in page_title:
                    print(f"    ✅ 在页面 {page_idx + 1} 标题中找到缩略语 '{acronym}'")
                    return page_idx
                
                # 检查文本块
                text_blocks = page.get("文本块", [])
                for text_block in text_blocks:
                    para_props = text_block.get("段落属性", [])
                    for para_prop in para_props:
                        content = para_prop.get("段落内容", "")
                        if acronym in content:
                            print(f"    ✅ 在页面 {page_idx + 1} 文本块中找到缩略语 '{acronym}'")
                            return page_idx
            
            print(f"    ❌ 未找到包含缩略语 '{acronym}' 的页面")
            return None
            
        except Exception as e:
            print(f"    ⚠️ 搜索缩略语页面时出错: {e}")
            return None
    
    def review_title_structure(self, parsing_data: Dict[str, Any]) -> List[Issue]:
        """审查标题结构：目录、章节、页面标题的层级一致性和逻辑连贯性"""
        print("    📋 审查标题结构...")
        # 提取页面内容
        pages = parsing_data.get("contents", [])
        
        # 使用提示词管理器获取用户提示词
        if self.prompt_manager:
            user_prompt = self.prompt_manager.get_user_prompt_for_review('title_structure')
            # 构建完整提示词：用户提示 + 输入提示 + 输出提示
            prompt = f"""{user_prompt}

                PPT内容：
                {json.dumps(pages, ensure_ascii=False, indent=2)}

                **重要**：请为每个问题提供精确的对象引用，格式如下：
                - 如果问题影响整个页面：使用 "page_[页码]"
                - 如果问题在特定文本块中：使用 "text_block_[页码]_[块索引]"
                - 如果问题在标题中：使用 "title_[页码]"

                请以JSON格式返回审查结果，格式如下：
                {{
                    "issues": [
                        {{
                            "rule_id": "LLM_TitleStructureRule",
                            "severity": "warning|info|serious",
                            "slide_index": 1（注意：页码从1开始计数）,
                            "object_ref": "title_1（注意：页码从1开始计数）,
                            "message": "问题描述",
                            "suggestion": "具体建议",
                            "can_autofix": false
                        }}
                    ]
                }}

                只返回JSON，不要其他内容。"""
        else:
            # 回退到默认提示词
            prompt = self._get_default_title_structure_prompt(pages)
        
        try:
            response = self.llm.complete(prompt, max_tokens=self.config.llm_max_tokens, stop_event=self.stop_event)
            if response:
                print(f"    📥 收到LLM响应，长度: {len(response)} 字符")
                print(f"    📄 响应前100字符: {response[:100]}...")
                
                cleaned_response = self._clean_json_response(response)
                print(f"    🧹 清理后响应长度: {len(cleaned_response)} 字符")
                
                # 尝试解析JSON
                try:
                    result = json.loads(cleaned_response)
                    print(f"    ✅ JSON解析成功")
                except json.JSONDecodeError as json_error:
                    print(f"    ❌ JSON解析失败: {json_error}")
                    print(f"    📄 清理后的响应内容:")
                    print(f"    {cleaned_response}")
                    
                    # 如果清理后的响应为空，显示原始响应
                    if not cleaned_response.strip():
                        print(f"    📄 原始响应内容:")
                        print(f"    {response}")
                        print(f"    ⚠️ LLM可能返回了空响应或非JSON格式的响应")
                    
                    # 尝试进一步修复
                    try:
                        # 查找可能的JSON部分
                        import re
                        json_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
                        matches = re.findall(json_pattern, cleaned_response)
                        if matches:
                            # 尝试解析找到的JSON部分
                            for i, match in enumerate(matches):
                                try:
                                    result = json.loads(match)
                                    print(f"    🔧 找到并解析JSON部分 {i+1}: {match[:100]}...")
                                    break
                                except:
                                    continue
                            else:
                                print(f"    ❌ 所有找到的JSON部分都无法解析")
                                return []
                        else:
                            print(f"    ❌ 未找到有效的JSON结构")
                            return []
                    except Exception as fix_error:
                        print(f"    ❌ JSON修复尝试失败: {fix_error}")
                        return []
                
                # 验证JSON结构
                if not isinstance(result, dict):
                    print(f"    ❌ 响应不是有效的JSON对象")
                    return []
                
                if "issues" not in result:
                    print(f"    ❌ 响应中缺少'issues'字段")
                    return []
                
                issues = []
                for i, item in enumerate(result.get("issues", [])):
                    try:
                        # 验证必要字段
                        if not isinstance(item, dict):
                            print(f"    ⚠️ 跳过无效的问题项 {i}: 不是字典类型")
                            continue
                        
                        # 处理页码：将LLM返回的页码（从1开始）转换为数组索引（从0开始）
                        slide_index = item.get("slide_index", 1)
                        if not isinstance(slide_index, (int, float)):
                            print(f"    ⚠️ 跳过问题项 {i}: slide_index不是数字类型")
                            continue
                        
                        array_index = max(0, int(slide_index) - 1)  # 确保不会小于0
                        
                        issue = Issue(
                            file="",
                            slide_index=array_index,  # 使用数组索引（从0开始）
                            object_ref=item.get("object_ref", "page"),
                            rule_id=item.get("rule_id", "LLM_TitleStructureRule"),
                            severity=item.get("severity", "info"),
                            message=item.get("message", ""),
                            suggestion=item.get("suggestion", ""),
                            can_autofix=item.get("can_autofix", False)
                        )
                        issues.append(issue)
                        print(f"    ✅ 添加问题: {issue.rule_id} - {issue.message[:50]}...")
                        
                    except Exception as item_error:
                        print(f"    ⚠️ 处理问题项 {i} 时出错: {item_error}")
                        continue
                
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
    
    def optimize_report(self, report_md: str) -> Optional[str]:
        """使用LLM优化报告：去重、精简内容"""
        if not report_md or not report_md.strip():
            return None
            
        # 使用提示词管理器获取用户提示词
        if self.prompt_manager:
            user_prompt = self.prompt_manager.get_user_prompt_for_review('report_optimization')
            # 构建完整提示词：用户提示 + 输入提示 + 输出提示
            prompt = f"""{user_prompt}

                **原始报告：**
                ```markdown
                {report_md}
                ```

                请返回优化后的报告，保持Markdown格式，确保：
                - 删除重复和冗余内容
                - 保留所有重要问题
                - 维持清晰的层次结构
                - 突出关键改进建议

                只返回优化后的Markdown报告，不要其他内容。"""
        else:
            # 回退到默认提示词
            prompt = self._get_default_report_optimization_prompt(report_md)
        
        try:
            print(f"    📤 发送报告优化请求...")
            print(f"    📝 原始报告长度: {len(report_md)} 字符")
            
            response = self.llm.complete(prompt, max_tokens=self.config.llm_max_tokens, stop_event=self.stop_event)
            
            if response and response.strip():
                # 清理响应，移除可能的markdown代码块标记
                optimized_report = response.strip()
                if optimized_report.startswith('```markdown'):
                    optimized_report = optimized_report[11:]
                elif optimized_report.startswith('```'):
                    optimized_report = optimized_report[3:]
                if optimized_report.endswith('```'):
                    optimized_report = optimized_report[:-3]
                optimized_report = optimized_report.strip()
                
                print(f"    📥 收到优化报告，长度: {len(optimized_report)} 字符")
                print(f"    📊 优化效果: 原始 {len(report_md)} → 优化后 {len(optimized_report)} 字符")
                
                return optimized_report
            else:
                print(f"    ⚠️ LLM未返回优化报告")
                return None
                
        except Exception as e:
            print(f"    ❌ 报告优化失败: {e}")
            return None

    def run_llm_review(self, doc: DocumentModel) -> List[Issue]:
        """运行完整的LLM审查流程"""
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
