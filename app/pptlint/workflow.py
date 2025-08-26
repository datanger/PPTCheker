"""
工作流编排器：根据《审查需求文档》与配置，动态串联组件。
支持规则+LLM混合审查模式，LLM不可用时自动降级为纯规则。
"""
from typing import List, Optional

from .config import ToolConfig
from .parser import parse_pptx
from .rules import run_basic_rules
from .reporter import render_markdown
from .annotator import annotate_pptx
from .model import Issue
from .llm import LLMClient
from .llm_review import create_llm_reviewer


class WorkflowResult:
    def __init__(self):
        self.issues: List[Issue] = []
        self.report_md: Optional[str] = None
        self.rule_issues_count: int = 0
        self.llm_issues_count: int = 0


def run_review_workflow(file_path: str, cfg: ToolConfig, output_ppt: Optional[str], llm: Optional[LLMClient]) -> WorkflowResult:
    res = WorkflowResult()
    
    # 步骤1：解析PPT
    print("📖 解析PPT文件...")
    doc = parse_pptx(file_path)
    
    # 步骤2：规则检查（基础格式检查）
    print("🔍 执行规则检查...")
    rule_issues = run_basic_rules(doc, cfg)
    res.rule_issues_count = len(rule_issues)
    print(f"✅ 规则检查完成，发现 {len(rule_issues)} 个问题")
    
    # 步骤3：LLM智能审查（如果可用）
    llm_issues = []
    if llm and llm.is_enabled():
        try:
            llm_reviewer = create_llm_reviewer(llm, cfg)
            llm_issues = llm_reviewer.run_llm_review(doc)
            res.llm_issues_count = len(llm_issues)
        except Exception as e:
            print(f"⚠️ LLM审查失败，降级为纯规则模式: {e}")
            llm_issues = []
    else:
        print("ℹ️ LLM未配置，使用纯规则审查模式")
    
    # 步骤4：合并所有问题
    all_issues = rule_issues + llm_issues
    res.issues = all_issues
    
    # 步骤5：生成报告
    print("📊 生成审查报告...")
    res.report_md = render_markdown(all_issues)
    
    # 步骤6：输出标记PPT（如果指定）
    if output_ppt:
        print("🏷️ 生成标记PPT...")
        annotate_pptx(file_path, all_issues, output_ppt)
    
    # 步骤7：输出统计信息
    print(f"\n🎯 审查完成！")
    print(f"   - 规则检查：{res.rule_issues_count} 个问题")
    print(f"   - LLM审查：{res.llm_issues_count} 个问题")
    print(f"   - 总计：{len(all_issues)} 个问题")
    
    return res


def run_edit_workflow(file_path: str, cfg: ToolConfig, output_ppt: str, llm: Optional[LLMClient]) -> WorkflowResult:
    """编辑模式：当前与审查模式相同，为未来自动修复预留接口"""
    print("✏️ 编辑模式：当前仅标记问题，自动修复功能待实现")
    return run_review_workflow(file_path, cfg, output_ppt, llm)


