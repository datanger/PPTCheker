"""
工作流工具模块：提供工作流所需的所有工具函数

功能：
1. 从 parsing_result.json 加载和解析数据
2. 执行基础规则检查
3. 执行LLM智能审查
4. 生成审查报告
5. 生成标记PPT

输入：parsing_result.json 格式的数据
输出：审查结果、报告、标记PPT等
"""

import json
import os
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path

# 导入必要的模块
try:
    from ..model import Issue, DocumentModel, Slide, Shape, TextRun, PPTContext, EditSuggestion, EditResult
    from ..config import ToolConfig
    from ..llm import LLMClient
    from ..reporter import render_markdown
    from ..annotator import annotate_pptx
except ImportError:
    # 兼容直接运行的情况
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from model import Issue, DocumentModel, Slide, Shape, TextRun
    from config import ToolConfig
    from llm import LLMClient
    from reporter import render_markdown
    from annotator import annotate_pptx


def load_parsing_result(file_path: str = "parsing_result.json") -> Dict[str, Any]:
    """加载 parsing_result.json 文件"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ 加载 parsing_result.json 失败: {e}")
        return {"页数": 0, "contents": []}


def convert_parsing_result_to_document_model(parsing_data: Dict[str, Any], file_path: str) -> DocumentModel:
    """将 parsing_result.json 格式转换为 DocumentModel 格式"""
    slides = []
    
    for page_data in parsing_data.get("contents", []):
        slide = Slide(
            index=page_data["页码"] - 1,  # 转换为0基索引
            slide_title=page_data.get("页标题", ""),
            slide_type=page_data.get("页类型", "内容页"),
            chapter_info=None
        )
        
        # 处理文本块（改：使用“段落属性”直接构建 TextRun，已不再依赖“拼接字符”）
        for text_block in page_data.get("文本块", []):
            shape = Shape(
                id=str(text_block["文本块索引"]),
                type="text",
                is_title=text_block.get("是否是标题占位符", False),
                title_level=1 if text_block.get("是否是标题占位符", False) else None,
                text_color=None,
                fill_color=None,
                border_color=None
            )
            
            # 直接从“段落属性”构建 TextRun
            para_runs = text_block.get("段落属性", [])
            for r in para_runs:
                tr = TextRun(
                    text=str(r.get("字符内容", "")),
                    font_name=r.get("字体类型"),
                    font_size_pt=float(r.get("字号")) if r.get("字号") is not None else None,
                    language_tag="ja",
                    is_bold=bool(r.get("是否粗体", False)),
                    is_italic=bool(r.get("是否斜体", False)),
                    is_underline=bool(r.get("是否下划线", False))
                )
                shape.text_runs.append(tr)
            
            slide.shapes.append(shape)
        
        slides.append(slide)
    
    doc = DocumentModel(file_path=file_path, slides=slides)
    return doc


def _parse_concatenated_text_to_runs(concatenated_text: str) -> List[TextRun]:
    """从拼接字符中解析出文本运行"""
    runs = []
    
    # 简单的文本分割（实际应该更智能地解析）
    # 这里简化处理，将整个文本作为一个run
    if concatenated_text:
        # 提取纯文本内容（去除属性标记）
        clean_text = _extract_clean_text(concatenated_text)
        
        run = TextRun(
            text=clean_text,
            font_name="默认字体",  # 从属性中提取
            font_size_pt=18.0,     # 从属性中提取
            language_tag="ja",      # 默认日语
            is_bold=False,
            is_italic=False,
            is_underline=False
        )
        runs.append(run)
    
    return runs


def _extract_clean_text(concatenated_text: str) -> str:
    """从拼接字符中提取纯文本内容"""
    # 移除所有属性标记
    import re
    
    # 移除【初始的字符所有属性：...】标记
    text = re.sub(r'【初始的字符所有属性：[^】]*】', '', concatenated_text)
    
    # 移除【字符属性变更：...】标记
    text = re.sub(r'【字符属性变更：[^】]*】', '', text)
    
    # 移除【换行】标记
    text = re.sub(r'【换行】', '\n', text)
    
    # 移除【缩进{...}】标记
    text = re.sub(r'【缩进\{\d+\}】', '', text)
    
    return text.strip()


def run_basic_rules(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    """运行基础规则检查"""
    from .rules import run_basic_rules as run_rules
    return run_rules(doc, cfg)


def run_llm_review(doc: DocumentModel, llm: LLMClient, cfg: ToolConfig) -> List[Issue]:
    """运行LLM智能审查"""
    try:
        from .llm_review import create_llm_reviewer
        reviewer = create_llm_reviewer(llm, cfg)
        return reviewer.run_llm_review(doc)
    except Exception as e:
        print(f"⚠️ LLM审查失败: {e}")
        return []


def generate_report(issues: List[Issue]) -> str:
    """生成审查报告"""
    try:
        return render_markdown(issues)
    except Exception as e:
        print(f"⚠️ 生成报告失败: {e}")
        return f"# 审查报告\n\n生成报告时发生错误: {e}"


def generate_annotated_ppt(input_ppt: str, issues: List[Issue], output_ppt: str) -> bool:
    """生成标记PPT"""
    try:
        annotate_pptx(input_ppt, issues, output_ppt)
        return True
    except Exception as e:
        print(f"⚠️ 生成标记PPT失败: {e}")
        return False


def get_workflow_statistics(rule_issues: List[Issue], llm_issues: List[Issue]) -> Dict[str, Any]:
    """获取工作流统计信息"""
    return {
        "rule_issues_count": len(rule_issues),
        "llm_issues_count": len(llm_issues),
        "total_issues": len(rule_issues) + len(llm_issues),
        "issues_by_severity": _count_issues_by_severity(rule_issues + llm_issues),
        "issues_by_rule": _count_issues_by_rule(rule_issues + llm_issues)
    }


# 新增：PPT编辑相关功能
def create_ppt_context(parsing_data: Dict[str, Any], original_pptx_path: str) -> Optional[PPTContext]:
    """创建PPT编辑上下文"""
    try:
        from pptx import Presentation
        
        # 加载原始PPT
        prs = Presentation(original_pptx_path)
        
        # 提取主题信息
        theme_info = extract_theme_info(prs)
        
        # 创建上下文对象
        context = PPTContext(
            parsing_result=parsing_data,
            original_pptx_path=original_pptx_path,
            presentation_object=prs,
            slide_layouts=list(prs.slide_layouts),
            slide_masters=list(prs.slide_masters),
            theme_info=theme_info
        )
        
        print(f"✅ 成功创建PPT编辑上下文，共 {len(prs.slides)} 页")
        return context
        
    except Exception as e:
        print(f"⚠️ 创建PPT上下文失败: {e}")
        return None


def extract_theme_info(prs) -> Dict[str, Any]:
    """提取PPT主题信息"""
    theme_info = {}
    try:
        # 提取主题色
        if hasattr(prs, 'core_properties'):
            theme_info['title'] = getattr(prs.core_properties, 'title', '')
            theme_info['author'] = getattr(prs.core_properties, 'author', '')
            theme_info['created'] = getattr(prs.core_properties, 'created', '')
        
        # 提取母版信息
        if hasattr(prs, 'slide_masters'):
            theme_info['slide_masters_count'] = len(prs.slide_masters)
        
        # 提取布局信息
        if hasattr(prs, 'slide_layouts'):
            theme_info['slide_layouts_count'] = len(prs.slide_layouts)
            
    except Exception as e:
        print(f"⚠️ 提取主题信息失败: {e}")
    
    return theme_info


def run_llm_edit_analysis(parsing_data: Dict[str, Any], llm: LLMClient, edit_requirements: str) -> List[EditSuggestion]:
    """使用LLM分析并生成编辑建议"""
    try:
        # 构建编辑分析提示词
        prompt = f"""
            你是PPT编辑专家。基于以下PPT内容分析，请提供具体的编辑建议：

            PPT内容：
            {json.dumps(parsing_data, ensure_ascii=False, indent=2)}

            编辑要求：
            {edit_requirements}

            请分析PPT内容，识别需要改进的地方，并输出JSON格式的编辑建议。

            输出格式（只输出JSON数组，不要解释）：
            [
            {{
                "type": "text_change|font_change|color_change|layout_change",
                "page_number": 1,
                "shape_index": 0,
                "current_value": "当前值",
                "new_value": "新值",
                "reason": "修改原因",
                "priority": "high|medium|low",
                "can_auto_apply": true
            }}
            ]

            注意：
            1. page_number 从1开始计数
            2. shape_index 是该页中形状的索引（从0开始）
            3. 只提供确实需要修改的建议
            4. 确保建议具体且可执行
            """
                    
        # 调用LLM
        response = llm.complete(prompt=prompt, max_tokens=2048)
        
        # 解析JSON响应
        try:
            suggestions_data = json.loads(response.strip())
            suggestions = []
            
            for item in suggestions_data:
                suggestion = EditSuggestion(
                    type=item.get('type', 'text_change'),
                    page_number=item.get('page_number', 1),
                    shape_index=item.get('shape_index', 0),
                    current_value=item.get('current_value', ''),
                    new_value=item.get('new_value', ''),
                    reason=item.get('reason', ''),
                    priority=item.get('priority', 'medium'),
                    can_auto_apply=item.get('can_auto_apply', True)
                )
                suggestions.append(suggestion)
            
            print(f"✅ LLM生成 {len(suggestions)} 个编辑建议")
            return suggestions
            
        except json.JSONDecodeError as e:
            print(f"⚠️ 解析LLM响应失败: {e}")
            print(f"LLM原始响应: {response}")
            return []
            
    except Exception as e:
        print(f"⚠️ LLM编辑分析失败: {e}")
        return []


def apply_edits_to_ppt(ppt_context: PPTContext, edit_suggestions: List[EditSuggestion]) -> EditResult:
    """应用编辑建议到PPT"""
    result = EditResult(success=False)
    
    if not ppt_context or not ppt_context.presentation_object:
        result.error_messages.append("PPT上下文无效")
        return result
    
    try:
        for suggestion in edit_suggestions:
            try:
                # 获取目标幻灯片
                slide = ppt_context.get_editable_slide(suggestion.page_number)
                if not slide:
                    result.failed_suggestions.append(suggestion)
                    result.error_messages.append(f"页面 {suggestion.page_number} 不存在")
                    continue
                
                # 获取目标形状
                if suggestion.shape_index >= len(slide.shapes):
                    result.failed_suggestions.append(suggestion)
                    result.error_messages.append(f"页面 {suggestion.page_number} 的形状 {suggestion.shape_index} 不存在")
                    continue
                
                shape = slide.shapes[suggestion.shape_index]
                
                # 根据类型应用编辑
                if suggestion.type == "text_change":
                    # 修改文本
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        shape.text_frame.text = suggestion.new_value
                        result.applied_suggestions.append(suggestion)
                        result.modified_slides.append(suggestion.page_number)
                        print(f"✅ 页面 {suggestion.page_number} 形状 {suggestion.shape_index} 文本已修改")
                    else:
                        result.failed_suggestions.append(suggestion)
                        result.error_messages.append(f"形状 {suggestion.shape_index} 不支持文本编辑")
                
                elif suggestion.type == "font_change":
                    # 修改字体
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.name = suggestion.new_value
                        result.applied_suggestions.append(suggestion)
                        result.modified_slides.append(suggestion.page_number)
                        print(f"✅ 页面 {suggestion.page_number} 形状 {suggestion.shape_index} 字体已修改")
                    else:
                        result.failed_suggestions.append(suggestion)
                        result.error_messages.append(f"形状 {suggestion.shape_index} 不支持字体编辑")
                
                elif suggestion.type == "color_change":
                    # 修改颜色
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        from ..model import Color
                        # 解析颜色值（假设格式为 #RRGGBB）
                        if suggestion.new_value.startswith('#'):
                            r = int(suggestion.new_value[1:3], 16)
                            g = int(suggestion.new_value[3:5], 16)
                            b = int(suggestion.new_value[5:7], 16)
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = (r, g, b)
                            result.applied_suggestions.append(suggestion)
                            result.modified_slides.append(suggestion.page_number)
                            print(f"✅ 页面 {suggestion.page_number} 形状 {suggestion.shape_index} 颜色已修改")
                        else:
                            result.failed_suggestions.append(suggestion)
                            result.error_messages.append(f"无效的颜色格式: {suggestion.new_value}")
                    else:
                        result.failed_suggestions.append(suggestion)
                        result.error_messages.append(f"形状 {suggestion.shape_index} 不支持颜色编辑")
                
                else:
                    result.failed_suggestions.append(suggestion)
                    result.error_messages.append(f"不支持的编辑类型: {suggestion.type}")
                
            except Exception as e:
                result.failed_suggestions.append(suggestion)
                result.error_messages.append(f"应用编辑建议失败: {e}")
                print(f"⚠️ 应用编辑建议失败: {e}")
        
        # 去重修改的页面
        result.modified_slides = list(set(result.modified_slides))
        result.success = len(result.applied_suggestions) > 0
        
        print(f"✅ 编辑完成：成功 {len(result.applied_suggestions)} 个，失败 {len(result.failed_suggestions)} 个")
        return result
        
    except Exception as e:
        result.error_messages.append(f"应用编辑过程中发生错误: {e}")
        print(f"⚠️ 应用编辑过程中发生错误: {e}")
        return result


def save_modified_ppt(ppt_context: PPTContext, output_path: str) -> bool:
    """保存修改后的PPT"""
    try:
        if ppt_context and ppt_context.presentation_object:
            ppt_context.presentation_object.save(output_path)
            print(f"✅ 修改后的PPT已保存到: {output_path}")
            return True
        else:
            print("❌ PPT上下文无效，无法保存")
            return False
    except Exception as e:
        print(f"⚠️ 保存PPT失败: {e}")
        return False


def _count_issues_by_severity(issues: List[Issue]) -> Dict[str, int]:
    """按严重程度统计问题"""
    counts = {"error": 0, "warning": 0, "info": 0}
    for issue in issues:
        severity = getattr(issue, 'severity', 'warning')
        counts[severity] = counts.get(severity, 0) + 1
    return counts


def _count_issues_by_rule(issues: List[Issue]) -> Dict[str, int]:
    """按规则类型统计问题"""
    counts = {}
    for issue in issues:
        rule_id = getattr(issue, 'rule_id', 'unknown')
        counts[rule_id] = counts.get(rule_id, 0) + 1
    return counts
