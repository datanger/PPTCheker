"""
核心数据模型（对应任务：实现PPTX解析构建DocumentModel 的数据承载）

设计目标：
- 提供解析层与规则引擎之间的稳定数据结构。
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


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
    # 新增：标题识别相关字段
    is_bold: Optional[bool] = None
    is_italic: Optional[bool] = None
    is_underline: Optional[bool] = None


@dataclass
class Shape:
    """形状抽象，统一承载文本与颜色。"""
    id: str
    type: str  # TextBox | TableCell | GraphicText 等
    text_runs: List[TextRun] = field(default_factory=list)
    text_color: Optional[Color] = None
    fill_color: Optional[Color] = None
    border_color: Optional[Color] = None
    # 新增：标题识别相关字段
    is_title: Optional[bool] = None
    title_level: Optional[int] = None  # 1=H1(主标题), 2=H2(章节标题), 3=H3(子标题)
    is_toc: Optional[bool] = None  # 是否为目录页面
    position: Optional[tuple] = None  # 位置信息 (left, top, width, height)


@dataclass
class Slide:
    index: int
    shapes: List[Shape] = field(default_factory=list)
    theme_colors: List[Color] = field(default_factory=list)
    # 新增：标题识别相关字段
    slide_title: Optional[str] = None  # 页面标题
    slide_type: Optional[str] = None  # 页面类型：title, content, toc, chapter等
    chapter_info: Optional[dict] = None  # 章节信息


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


# 新增：PPT编辑相关模型
@dataclass
class EditSuggestion:
    """编辑建议实体"""
    type: str  # text_change, font_change, color_change, layout_change
    page_number: int
    shape_index: int
    current_value: str
    new_value: str
    reason: str
    priority: str = "medium"  # high, medium, low
    can_auto_apply: bool = True


@dataclass
class PPTContext:
    """PPT编辑上下文，包含所有必要信息"""
    parsing_result: Dict[str, Any]           # 解析结果
    original_pptx_path: str                  # 原始PPT文件路径
    presentation_object: Optional[Any] = None  # python-pptx对象
    slide_layouts: List[Any] = field(default_factory=list)  # 幻灯片布局
    slide_masters: List[Any] = field(default_factory=list)  # 母版信息
    theme_info: Dict[str, Any] = field(default_factory=dict)  # 主题信息
    
    def get_editable_slide(self, page_number: int):
        """获取可编辑的幻灯片对象"""
        if self.presentation_object and 1 <= page_number <= len(self.presentation_object.slides):
            return self.presentation_object.slides[page_number - 1]
        return None
    
    def get_slide_layout(self, layout_index: int):
        """获取幻灯片布局"""
        if 0 <= layout_index < len(self.slide_layouts):
            return self.slide_layouts[layout_index]
        return None
    
    def get_slide_master(self, master_index: int):
        """获取幻灯片母版"""
        if 0 <= master_index < len(self.slide_masters):
            return self.slide_masters[master_index]
        return None


@dataclass
class EditResult:
    """编辑结果"""
    success: bool
    modified_slides: List[int] = field(default_factory=list)
    applied_suggestions: List[EditSuggestion] = field(default_factory=list)
    failed_suggestions: List[EditSuggestion] = field(default_factory=list)
    error_messages: List[str] = field(default_factory=list)
    output_path: Optional[str] = None

