"""
核心数据模型（对应任务：实现PPTX解析构建DocumentModel 的数据承载）

设计目标：
- 提供解析层与规则引擎之间的稳定数据结构。
"""
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Color:
    """颜色模型，简单采用 sRGB 三元组。"""
    r: int
    g: int
    b: int


@dataclass
class TextRun:
    """文本片段，承载字体/字号/语言标签等。
    - 对应需求：字体与字号检查、日文检测、术语一致性。
    """
    text: str
    font_name: Optional[str]
    font_size_pt: Optional[float]
    language_tag: Optional[str] = None


@dataclass
class Shape:
    """形状抽象，统一承载文本与颜色。"""
    id: str
    type: str  # TextBox | TableCell | GraphicText 等
    text_runs: List[TextRun] = field(default_factory=list)
    text_color: Optional[Color] = None
    fill_color: Optional[Color] = None
    border_color: Optional[Color] = None


@dataclass
class Slide:
    index: int
    shapes: List[Shape] = field(default_factory=list)
    theme_colors: List[Color] = field(default_factory=list)


@dataclass
class DocumentModel:
    file_path: str
    slides: List[Slide]


@dataclass
class Issue:
    """规则问题实体（报告项）。
    - 对应需求：报告输出字段。
    """
    file: str
    slide_index: int
    object_ref: str
    rule_id: str
    severity: str
    message: str
    suggestion: Optional[str] = None
    can_autofix: bool = False
    fixed: bool = False

