"""
工作流编排器：根据《审查需求文档》与配置，动态串联组件。
默认启用LLM（若配置齐全），否则降级为纯规则。
"""
from typing import List, Optional

from .config import ToolConfig
from .parser import parse_pptx
from .rules import run_basic_rules
from .reporter import render_markdown
from .annotator import annotate_pptx
from .model import Issue
from .llm import LLMClient


class WorkflowResult:
    def __init__(self):
        self.issues: List[Issue] = []
        self.report_md: Optional[str] = None


def run_review_workflow(file_path: str, cfg: ToolConfig, output_ppt: Optional[str], llm: Optional[LLMClient]) -> WorkflowResult:
    res = WorkflowResult()
    doc = parse_pptx(file_path)
    issues = run_basic_rules(doc, cfg)
    # 预留：如 llm.is_enabled()，在此追加语言/逻辑建议型 Issue
    res.issues = issues
    res.report_md = render_markdown(issues)
    if output_ppt:
        annotate_pptx(file_path, issues, output_ppt)
    return res


def run_edit_workflow(file_path: str, cfg: ToolConfig, output_ppt: str, llm: Optional[LLMClient]) -> WorkflowResult:
    # 简化：先复用 review 流，后续在此对安全项直接改写
    return run_review_workflow(file_path, cfg, output_ppt, llm)


