"""
基础规则实现（对应任务：实现基础规则：字体字号、缩略语、色调一致、颜色数量）
"""
import re
from collections import defaultdict
from typing import List

from .model import DocumentModel, Issue, Color
from .config import ToolConfig


def _is_color_equal(c1: Color, c2: Color, tol: int = 10) -> bool:
    """近似颜色比较，阈值默认10/255。"""
    return (
        abs(c1.r - c2.r) <= tol and
        abs(c1.g - c2.g) <= tol and
        abs(c1.b - c2.b) <= tol
    )


def check_font_and_size(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
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
                # 日文字体检查
                if tr.language_tag == "ja":
                    font_name_norm = (tr.font_name or "").strip()
                    if font_name_norm != cfg.jp_font_name:
                        issues.append(Issue(
                            file=doc.file_path,
                            slide_index=slide.index,
                            object_ref=shp.id,
                            rule_id="FontFamilyRule",
                            severity="warning",
                            message=f"日文字体非 {cfg.jp_font_name}: {font_name_norm or '未指定'}",
                            suggestion=f"替换为 {cfg.jp_font_name}",
                            can_autofix=cfg.autofix_font,
                        ))
    return issues


_ACRONYM_RE = re.compile(r"\b([A-Z]{2,8})\b")


def check_acronym_explanation(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    issues: List[Issue] = []
    seen = set()
    for slide in doc.slides:
        slide_text = " ".join(tr.text for shp in slide.shapes for tr in shp.text_runs)
        for m in _ACRONYM_RE.finditer(slide_text):
            ac = m.group(1)
            if len(ac) < cfg.acronym_min_len or len(ac) > cfg.acronym_max_len:
                continue
            if ac not in seen:
                # 简单策略：同页是否存在括号或冒号解释
                if not re.search(rf"\b{ac}\b\s*(\(|：|:)", slide_text):
                    issues.append(Issue(
                        file=doc.file_path,
                        slide_index=slide.index,
                        object_ref="page",
                        rule_id="AcronymRule",
                        severity="info",
                        message=f"缩略语 {ac} 首次出现未发现解释",
                        suggestion=f"在首次出现后添加解释：{ac}: <全称>",
                        can_autofix=False,
                    ))
                seen.add(ac)
    return issues


def check_color_count(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
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
    # 简化：由于主题色提取依赖更深入的母版解析，这里先留空返回[]，后续演进
    return []


def run_basic_rules(doc: DocumentModel, cfg: ToolConfig) -> List[Issue]:
    issues: List[Issue] = []
    issues += check_font_and_size(doc, cfg)
    issues += check_acronym_explanation(doc, cfg)
    issues += check_color_count(doc, cfg)
    issues += check_theme_harmony(doc, cfg)
    return issues

