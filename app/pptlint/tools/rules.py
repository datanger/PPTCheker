"""
基础规则实现（对应任务：实现基础规则：字体字号、颜色数量等明确格式检查）
删除涉及语义理解的规则，这些交给大模型处理
"""
import re
from collections import defaultdict
from typing import List

try:
    from ..model import DocumentModel, Issue, Color
    from ..config import ToolConfig
except ImportError:
    # 兼容直接运行的情况
    import sys
    import os
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from model import DocumentModel, Issue, Color
    from config import ToolConfig


def _is_color_equal(c1: Color, c2: Color, tol: int = 10) -> bool:
    """近似颜色比较，阈值默认10/255。"""
    return (
        abs(c1.r - c2.r) <= tol and
        abs(c1.g - c2.g) <= tol and
        abs(c1.b - c2.b) <= tol
    )


def check_font_and_size(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    """检查字体和字号（明确的格式规范）"""
    issues: List[Issue] = []
    for slide in doc.slides:
        for shp in slide.shapes:
            for tr in shp.text_runs:
                # 字号检查
                if tr.font_size_pt is not None and tr.font_size_pt < cfg.min_font_size_pt:
                    issues.append(Issue(
                        file=doc.file_path,
                        slide_index=slide.index,
                        object_ref=shp.id,
                        rule_id="FontSizeRule",
                        severity="warning",
                        message=f"字号 {tr.font_size_pt} < {cfg.min_font_size_pt}",
                        suggestion=f"提升至 {cfg.min_font_size_pt}pt",
                        can_autofix=cfg.autofix_size,
                    ))
                # 日文字体检查 - 过滤掉"未知"字体，只检查识别到的字体
                if tr.language_tag == "ja":
                    font_name_norm = (tr.font_name or "").strip()
                    # 只对识别到的字体进行检查，跳过"未知"字体
                    if font_name_norm and font_name_norm != "未知" and font_name_norm != cfg.jp_font_name:
                        issues.append(Issue(
                            file=doc.file_path,
                            slide_index=slide.index,
                            object_ref=shp.id,
                            rule_id="FontFamilyRule",
                            severity="warning",
                            message=f"日文字体非 {cfg.jp_font_name}: {font_name_norm}",
                            suggestion=f"替换为 {cfg.jp_font_name}",
                            can_autofix=cfg.autofix_font,
                        ))
    return issues


def check_color_count(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    """检查颜色数量（明确的格式规范）"""
    issues: List[Issue] = []
    for slide in doc.slides:
        color_set = set()
        def add(c: Color):
            if c is not None:
                color_set.add((c.r, c.g, c.b))
        for shp in slide.shapes:
            add(shp.text_color)
            add(shp.fill_color)
            add(shp.border_color)
        if len(color_set) > cfg.color_count_threshold:
            issues.append(Issue(
                file=doc.file_path,
                slide_index=slide.index,
                object_ref="page",
                rule_id="ColorCountRule",
                severity="warning",
                message=f"单页颜色数 {len(color_set)} 超过阈值 {cfg.color_count_threshold}",
                suggestion="减少临时色，统一为主题色",
                can_autofix=False,
            ))
    return issues


def check_theme_harmony(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    """检查主题色调一致性（预留接口，待完善）"""
    # 简化：由于主题色提取依赖更深入的母版解析，这里先留空返回[]，后续演进
    return []


def run_basic_rules(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    """运行基础规则（只包含明确的格式检查）"""
    issues: List[Issue] = []
    issues += check_font_and_size(doc, cfg)
    issues += check_color_count(doc, cfg)
    issues += check_theme_harmony(doc, cfg)
    return issues

