"""
PPTX 解析器 - 提取PPT文件的详细信息

输出格式：
- 每页为一个对象映射：{"文本块N": {...}, "图片M": {...}}
- 文本块：文本块位置、图层编号、是否是标题占位符、字符属性数组、拼接字符
- 图片：图片位置、类型、大小、图层位置

注意：某些字段可能无法直接获取，已删除并说明原因
"""

import re
import json
from typing import List, Dict, Any, Optional
from pptx import Presentation
from pptx.dml.color import RGBColor, MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
from pptx.util import Inches

# 兼容脚本直跑与包内相对导入
try:
    from .serializer import serialize_text_block_to_diff_string
except Exception:
    from app.pptlint.serializer import serialize_text_block_to_diff_string

# 删除的字段及原因：
# - 图片质量：无法直接获取，需要图像分析
# - 图片格式：只能获取文件扩展名，无法获取实际格式

# 主题色近似映射（Office 默认主题近似值）
THEME_COLOR_TO_HEX = {
    MSO_THEME_COLOR.TEXT_1: "#000000",
    MSO_THEME_COLOR.BACKGROUND_1: "#FFFFFF",
    MSO_THEME_COLOR.TEXT_2: "#44546A",
    MSO_THEME_COLOR.BACKGROUND_2: "#E7E6E6",
    MSO_THEME_COLOR.ACCENT_1: "#5B9BD5",
    MSO_THEME_COLOR.ACCENT_2: "#ED7D31",
    MSO_THEME_COLOR.ACCENT_3: "#A5A5A5",
    MSO_THEME_COLOR.ACCENT_4: "#FFC000",
    MSO_THEME_COLOR.ACCENT_5: "#4472C4",
    MSO_THEME_COLOR.ACCENT_6: "#70AD47",
    MSO_THEME_COLOR.HYPERLINK: "#0563C1",
    MSO_THEME_COLOR.FOLLOWED_HYPERLINK: "#954F72",
}


def _rgb_to_hex(color) -> Optional[str]:
    """将RGB/主题颜色转换为十六进制格式"""
    try:
        if isinstance(color, RGBColor):
            return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"
        if hasattr(color, "rgb") and color.rgb is not None:
            rgb = color.rgb
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        if hasattr(color, "theme_color") and color.theme_color is not None:
            return THEME_COLOR_TO_HEX.get(color.theme_color, None)
    except Exception:
        pass
    return None


def _get_shape_position(shape) -> Dict[str, float]:
    try:
        return {
            "left": float(shape.left),
            "top": float(shape.top),
            "width": float(shape.width),
            "height": float(shape.height)
        }
    except Exception:
        return {"left": 0.0, "top": 0.0, "width": 0.0, "height": 0.0}


def _is_title_placeholder(shape) -> bool:
    try:
        if shape.is_placeholder and (shape.placeholder_format.type == PP_PLACEHOLDER.TITLE or
                                     shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE or
                                     shape.placeholder_format.type == PP_PLACEHOLDER.VERTICAL_TITLE or
                                     shape.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE):
            return True
    except Exception:
        pass
    return False


def _resolve_run_font_props(run, para) -> Dict[str, Any]:
    props: Dict[str, Any] = {}
    try:
        rf = run.font
        pf = getattr(para, 'font', None)
        name = getattr(rf, 'name', None)
        if name is None and pf is not None:
            name = getattr(pf, 'name', None)
        if name is None:
            try:
                rPr = run._r.rPr
                if rPr is not None and getattr(rPr, 'rFonts', None) is not None:
                    rfonts = rPr.rFonts
                    name = getattr(rfonts, 'eastAsia', None) or getattr(rfonts, 'ascii', None) or getattr(rfonts, 'cs', None)
                if name is None and rPr is not None and getattr(rPr, 'latin', None) is not None:
                    name = rPr.latin.typeface
            except Exception:
                pass
        if name is None:
            name = "默认"
        props["字体类型"] = name
        size = getattr(rf, 'size', None)
        if size is None and pf is not None:
            size = getattr(pf, 'size', None)
        props["字号"] = float(size.pt) if size is not None else None
        color = getattr(rf, 'color', None)
        hex_color = _rgb_to_hex(color) if color is not None else None
        if hex_color is None and pf is not None:
            pcolor = getattr(pf, 'color', None)
            hex_color = _rgb_to_hex(pcolor) if pcolor is not None else None
        if hex_color is None:
            hex_color = "#000000"
        props["字体颜色"] = hex_color
        bold = getattr(rf, 'bold', None)
        if bold is None and pf is not None:
            bold = getattr(pf, 'bold', None)
        italic = getattr(rf, 'italic', None)
        if italic is None and pf is not None:
            italic = getattr(pf, 'italic', None)
        underline = getattr(rf, 'underline', None)
        if underline is None and pf is not None:
            underline = getattr(pf, 'underline', None)
        strike = getattr(rf, 'strike', None)
        if strike is None and pf is not None:
            strike = getattr(pf, 'strike', None)
        props.update({
            "是否粗体": bool(bold) if bold is not None else False,
            "是否斜体": bool(italic) if italic is not None else False,
            "是否下划线": bool(underline) if underline is not None else False,
            "是否带删除线": bool(strike) if strike is not None else False,
        })
    except Exception:
        pass
    return props


def _get_text_block_info(shape, shape_index: int) -> Dict[str, Any]:
    text_info = {}
    try:
        if shape.has_text_frame and shape.text_frame:
            text_block_position = _get_shape_position(shape)
            is_title_placeholder = _is_title_placeholder(shape)
            character_attributes = []
            character_index = 0
            paragraphs = shape.text_frame.paragraphs
            total_paras = len(paragraphs)
            for p_idx, para in enumerate(paragraphs):
                for run in para.runs:
                    eff = _resolve_run_font_props(run, para)
                    for char in run.text:
                        char_info = {
                            "字符编号": character_index,
                            "字符内容": char,
                            "字体类型": eff.get("字体类型"),
                            "字号": eff.get("字号"),
                            "字体颜色": eff.get("字体颜色"),
                            "是否粗体": eff.get("是否粗体"),
                            "是否斜体": eff.get("是否斜体"),
                            "是否下划线": eff.get("是否下划线"),
                            "是否带删除线": eff.get("是否带删除线"),
                        }
                        char_info = {k: v for k, v in char_info.items() if v is not None}
                        character_attributes.append(char_info)
                        character_index += 1
                if p_idx < total_paras - 1:
                    character_attributes.append({
                        "字符编号": character_index,
                        "字符内容": "\n"
                    })
                    character_index += 1
            if character_attributes:
                text_key = f"文本块{shape_index + 1}"
                text_payload = {
                    "文本块位置": text_block_position,
                    "图层编号": shape_index,
                    "是否是标题占位符": is_title_placeholder,
                    "字符属性": character_attributes
                }
                # 生成拼接字符
                try:
                    text_payload["拼接字符"] = serialize_text_block_to_diff_string({text_key: text_payload}, initial_label="初始的字符所有属性")
                except Exception:
                    text_payload["拼接字符"] = ""
                text_info = {text_key: text_payload}
    except Exception as e:
        print(f"提取文本块信息失败: {e}")
    return text_info


def _get_image_info(shape, shape_index: int) -> Dict[str, Any]:
    image_info = {}
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            position = _get_shape_position(shape)
            image_type = "未知"
            try:
                if hasattr(shape, 'image') and shape.image:
                    filename = getattr(shape.image, 'filename', '')
                    if filename:
                        image_type = filename.split('.')[-1].upper() if '.' in filename else "未知"
            except Exception:
                pass
            image_size = 0
            try:
                if hasattr(shape, 'image') and shape.image:
                    image_size = len(shape.image.blob) if hasattr(shape.image, 'blob') else 0
            except Exception:
                pass
            image_info = {
                f"图片{shape_index + 1}": {
                    "图片位置": position,
                    "图片类型": image_type,
                    "图片大小": f"{image_size} bytes",
                    "图层位置": shape_index
                }
            }
    except Exception as e:
        print(f"提取图片信息失败: {e}")
    return image_info


def parse_pptx(path: str) -> List[Dict[str, Any]]:
    try:
        prs = Presentation(path)
        slides_data: List[Dict[str, Any]] = []
        for slide_index, slide in enumerate(prs.slides):
            page_map: Dict[str, Any] = {}
            sorted_shapes = sorted(slide.shapes, key=lambda s: (s.top, s.left))
            for shape_index, shape in enumerate(sorted_shapes):
                text_info = _get_text_block_info(shape, shape_index)
                if text_info:
                    page_map.update(text_info)
                image_info = _get_image_info(shape, shape_index)
                if image_info:
                    page_map.update(image_info)
            slides_data.append(page_map)
        return slides_data
    except Exception as e:
        print(f"解析PPTX文件失败: {e}")
        return []


def save_to_json(data: List[Dict[str, Any]], output_path: str):
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"数据已保存到: {output_path}")
    except Exception as e:
        print(f"保存JSON文件失败: {e}")


if __name__ == "__main__":
    pptx_path = "example1.pptx"
    print("🧪 测试PPTX解析...")
    result = parse_pptx(pptx_path)
    if result:
        print(f"✅ 成功解析，共 {len(result)} 页")
        save_to_json(result, "parsing_result_new.json")
        first_page = result[0]
        print(f"\n📄 第一页包含 {len(first_page)} 个元素")
        # 显示第一个文本块的拼接字符片段
        for k, v in first_page.items():
            if k.startswith("文本块"):
                s = v.get("拼接字符", "")
                print(s[:200])
                break
    else:
        print("❌ 解析失败")

