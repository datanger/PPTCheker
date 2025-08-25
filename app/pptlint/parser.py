"""
PPTX 解析器（对应任务：实现PPTX解析构建DocumentModel（文本与颜色））

说明：
- 使用 python-pptx 提取幻灯片、形状、文本片段与基础颜色；
- 颜色近似按 sRGB 三元组记录；
- 语言标签简化：通过字符范围判断是否包含日文（平假名/片假名/日文汉字）。
"""
from typing import List, Optional
from pptx import Presentation
from pptx.dml.color import RGBColor

from .model import DocumentModel, Slide, Shape, TextRun, Color


def _rgb_of(color) -> Optional[Color]:
    """从 python-pptx 颜色对象读取 sRGB。"""
    try:
        if isinstance(color, RGBColor):
            return Color(r=color[0], g=color[1], b=color[2])
        if hasattr(color, "rgb") and color.rgb is not None:
            rgb: RGBColor = color.rgb
            return Color(r=rgb[0], g=rgb[1], b=rgb[2])
    except Exception:
        return None
    return None


def _detect_language_tag(text: str) -> Optional[str]:
    """极简语言检测：若包含日文字符，标注为 'ja'。"""
    for ch in text:
        code = ord(ch)
        # 平假名、片假名、CJK统一表意（粗略）
        if (0x3040 <= code <= 0x30FF) or (0x31F0 <= code <= 0x31FF) or (0x4E00 <= code <= 0x9FFF):
            return "ja"
    return None


def parse_pptx(path: str) -> DocumentModel:
    prs = Presentation(path)
    slides: List[Slide] = []

    for s_idx, slide in enumerate(prs.slides):
        shapes: List[Shape] = []
        for shp in slide.shapes:
            # 仅处理具有文本框架的形状
            if not hasattr(shp, "has_text_frame") or not shp.has_text_frame:
                continue
            text_runs: List[TextRun] = []
            text_color: Optional[Color] = None
            fill_color: Optional[Color] = None
            border_color: Optional[Color] = None

            try:
                if shp.text_frame is not None:
                    for p in shp.text_frame.paragraphs:
                        for run in p.runs:
                            txt = run.text or ""
                            font = run.font
                            font_name = font.name if font is not None else None
                            font_size_pt = float(font.size.pt) if (font is not None and font.size) else None
                            language_tag = _detect_language_tag(txt)
                            text_runs.append(
                                TextRun(text=txt, font_name=font_name, font_size_pt=font_size_pt, language_tag=language_tag)
                            )
                            if font is not None and font.color is not None:
                                tc = _rgb_of(font.color)
                                text_color = tc or text_color
            except Exception:
                pass

            try:
                if hasattr(shp, "fill") and shp.fill is not None and hasattr(shp.fill, "fore_color"):
                    fill_color = _rgb_of(shp.fill.fore_color) or fill_color
            except Exception:
                pass

            try:
                if hasattr(shp, "line") and shp.line is not None and hasattr(shp.line, "color"):
                    border_color = _rgb_of(shp.line.color) or border_color
            except Exception:
                pass

            shape = Shape(
                id=str(getattr(shp, "shape_id", f"s{len(shapes)}")),
                type=type(shp).__name__,
                text_runs=text_runs,
                text_color=text_color,
                fill_color=fill_color,
                border_color=border_color,
            )
            if text_runs:
                shapes.append(shape)

        slides.append(Slide(index=s_idx, shapes=shapes, theme_colors=[]))

    return DocumentModel(file_path=path, slides=slides)

