"""
PPTX 解析器 - 提取PPT文件的详细信息

输出格式：
- 每页为一个对象：{"页码": int, "文本块": [...], "图片": [...]}
- 文本块：文本块位置、图层编号、是否是标题占位符、拼接字符
- 图片：图片位置、类型、大小、图层位置

注意：某些字段可能无法直接获取，已删除并说明原因
"""

import json
from typing import List, Dict, Any, Optional, Tuple
from pptx import Presentation
from pptx.dml.color import RGBColor, MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
# 本文件不再生成“拼接字符”，改为直接输出“段落属性”（按 run 维度）

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


THEME_PLACEHOLDER_DEFAULT_MAP = {
    "+mn-ea": "微软雅黑",
    "+mj-ea": "微软雅黑",
    "+mn-lt": "Calibri",
    "+mj-lt": "Calibri",
}

# 用于在JSON层面合并相邻run：比较除“段落内容”外的样式是否一致
ATTR_COMPARE_KEYS = [
    "字体类型",
    "字号",
    "字体颜色",
    "是否粗体",
    "是否斜体",
    "是否下划线",
    "是否带删除线",
]


def _hex_to_rgb_tuple(hex_str: str) -> Tuple[int, int, int]:
    hex_str = hex_str.lstrip('#')
    return int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)


def _rgb_tuple_to_hex(rgb: Tuple[int, int, int]) -> str:
    r, g, b = rgb
    r = max(0, min(255, r))
    g = max(0, min(255, g))
    b = max(0, min(255, b))
    return f"#{r:02X}{g:02X}{b:02X}"


def _apply_brightness(base_hex: str, brightness: Optional[float]) -> str:
    """对主题色应用亮度调整（-1.0~1.0），仿照PowerPoint tint/shade 逻辑的近似实现。"""
    if brightness is None or brightness == 0:
        return base_hex
    r, g, b = _hex_to_rgb_tuple(base_hex)
    if brightness > 0:
        # tint toward white
        r = r + (255 - r) * brightness
        g = g + (255 - g) * brightness
        b = b + (255 - b) * brightness
    else:
        # shade toward black
        factor = 1.0 + brightness  # brightness is negative
        r = r * factor
        g = g * factor
        b = b * factor
    return _rgb_tuple_to_hex((int(round(r)), int(round(g)), int(round(b))))


def _rgb_to_hex(color) -> Optional[str]:
    """将RGB/主题颜色转换为十六进制格式，并考虑亮度调整。"""
    try:
        # 明确RGB
        if isinstance(color, RGBColor):
            return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"
        # python-pptx Font.color 可能有 rgb / theme_color / brightness
        if hasattr(color, "rgb") and color.rgb is not None:
            rgb = color.rgb
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        base_hex = None
        if hasattr(color, "theme_color") and color.theme_color is not None:
            base_hex = THEME_COLOR_TO_HEX.get(color.theme_color, None)
        if base_hex is not None:
            # brightness: -1.0..1.0（正值提亮，负值压暗）
            bright = getattr(color, "brightness", None)
            try:
                if bright is not None:
                    base_hex = _apply_brightness(base_hex, float(bright))
            except Exception:
                pass
            return base_hex
    except Exception:
        pass
    return None


# 将常见颜色的十六进制值映射为中文颜色名；其余按“灰色”近似
COMMON_COLOR_NAME_TO_RGB = {
    "黑色": (0, 0, 0),
    "白色": (255, 255, 255),
    "红色": (255, 0, 0),
    "绿色": (0, 255, 0),
    "蓝色": (0, 0, 255),
    "黄色": (255, 255, 0),
    "橙色": (255, 165, 0),
    "紫色": (128, 0, 128),
    "灰色": (128, 128, 128),
}


def _hex_to_cn_color_name(hex_color: str) -> str:
    """将 #RRGGBB 映射为最接近的常见中文颜色名。"""
    try:
        r, g, b = _hex_to_rgb_tuple(hex_color)
        best_name = "灰色"
        best_dist = float("inf")
        for name, (cr, cg, cb) in COMMON_COLOR_NAME_TO_RGB.items():
            dist = (r - cr) ** 2 + (g - cg) ** 2 + (b - cb) ** 2
            if dist < best_dist:
                best_dist = dist
                best_name = name
        return best_name
    except Exception:
        return "灰色"


def _merge_font_family_alias(raw_name: Optional[str]) -> str:
    """合并常见字体族别名/派生名到主名。
    规则示例：
    - "宋体"、"宋体-正文"、"宋体-标题" → "宋体"
    - "Meiryo"、"Meiryo-正文"、"Meiryo-Regular" → "Meiryo"
    其它字体保持原样；空值返回 "未知"。
    """
    if not isinstance(raw_name, str) or not raw_name.strip():
        return "未知"
    name = raw_name.strip()
    low = name.lower()
    # 去掉常见的后缀标记
    strip_suffixes = ["-正文", "-标题", "-regular", " regular", " bold", "-bold", " italic", "-italic"]
    for suf in strip_suffixes:
        if low.endswith(suf):
            name = name[: len(name) - len(suf)]
            low = name.lower()
            break
    # 统一 Meiryo 派生
    if "meiryo" in low:
        return "Meiryo"
    # 统一 宋体 派生
    if "宋体" in name:
        return "宋体"
    return name


def _get_shape_position(shape) -> Dict[str, str]:
    """返回形状位置，单位为百分比（%），相对左上角。
    PowerPoint内部单位为EMU，需要先获取幻灯片尺寸来计算百分比。
    """
    try:
        # 获取幻灯片尺寸 - 尝试多种方法
        slide_width = None
        slide_height = None
        
        try:
            # 方法1：直接从shape.part获取
            slide_width = shape.part.slide_width
            slide_height = shape.part.slide_height
        except Exception:
            pass
        
        try:
            # 方法2：从slide对象获取
            if hasattr(shape, 'slide'):
                slide_width = shape.slide.slide_width
                slide_height = shape.slide.slide_height
        except Exception:
            pass
        
        try:
            # 方法3：从slide_layout获取
            if hasattr(shape.part, 'slide_layout'):
                slide_width = shape.part.slide_layout.slide_width
                slide_height = shape.part.slide_layout.slide_height
        except Exception:
            pass
        
        # 如果仍然无法获取，使用默认值（标准PPT尺寸：16:9 宽屏 = 9144000 x 6858000 EMU）
        if slide_width is None or slide_height is None:
            slide_width = 9144000  # 16:9 宽屏宽度 (10" x 914400 EMU/inch)
            slide_height = 6858000  # 16:9 宽屏高度 (7.5" x 914400 EMU/inch)
        
        # 调试信息：打印实际获取的尺寸
        if hasattr(shape, '_element') and hasattr(shape._element, 'attrib'):
            try:
                # 尝试从XML属性获取实际尺寸
                xml_attrib = shape._element.attrib
                if 'cx' in xml_attrib and 'cy' in xml_attrib:
                    actual_width = int(xml_attrib['cx'])
                    actual_height = int(xml_attrib['cy'])
                    # 如果XML中的尺寸更合理，使用它
                    if actual_width > 0 and actual_height > 0:
                        slide_width = actual_width
                        slide_height = actual_height
            except Exception:
                pass
        
        def emu_to_percent_str(emu_val, slide_dimension) -> str:
            try:
                percent = (float(emu_val) / float(slide_dimension)) * 100.0
                # 不限制百分比范围，显示真实值（可能超过100%）
                return f"{percent:.2f}%"
            except Exception:
                return "0.00%"
        
        # 添加调试信息
        debug_info = {
            "slide_width_emu": slide_width,
            "slide_height_emu": slide_height,
            "shape_left_emu": shape.left,
            "shape_top_emu": shape.top,
            "shape_width_emu": shape.width,
            "shape_height_emu": shape.height
        }
        
        # 计算实际百分比（不限制范围，用于调试）
        actual_percentages = {
            "left": (float(shape.left) / float(slide_width)) * 100.0,
            "top": (float(shape.top) / float(slide_height)) * 100.0,
            "width": (float(shape.width) / float(slide_width)) * 100.0,
            "height": (float(shape.height) / float(slide_height)) * 100.0
        }
        
        return {
            "left": emu_to_percent_str(shape.left, slide_width),
            "top": emu_to_percent_str(shape.top, slide_height),
            "width": emu_to_percent_str(shape.width, slide_width),
            "height": emu_to_percent_str(shape.height, slide_height)
        }
    except Exception as e:
        print(f"获取形状位置失败: {e}")
        return {"left": "0.00%", "top": "0.00%", "width": "0.00%", "height": "0.00%"}


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


def _get_para_rfonts(para) -> Optional[str]:
    """尝试从段落 pPr.rPr.rFonts / latin 中获取字体名称。"""
    try:
        p = getattr(para, '_p', None)
        pPr = getattr(p, 'pPr', None) if p is not None else None
        rPr = getattr(pPr, 'rPr', None) if pPr is not None else None
        if rPr is not None:
            # rFonts
            rfonts = getattr(rPr, 'rFonts', None)
            if rfonts is not None:
                name = (
                    getattr(rfonts, 'eastAsia', None)
                    or getattr(rfonts, 'ascii', None)
                    or getattr(rfonts, 'hAnsi', None)
                    or getattr(rfonts, 'cs', None)
                )
                if name:
                    return name
            # latin
            latin = getattr(rPr, 'latin', None)
            if latin is not None and getattr(latin, 'typeface', None):
                return latin.typeface
    except Exception:
        pass
    return None


def _resolve_run_font_props(run, para, is_title_placeholder: bool, host_shape) -> Dict[str, Any]:
    """解析 run 的样式属性。
    要求：字符类型（字体名称）仅保留 run/paragraph 级别识别：
    1) run._r.rPr.rFonts/latin
    2) run.font.name
    3) para.font.name
    4) para 的 pPr.rPr.rFonts/latin
    不再使用占位符、shape.lstStyle、母版textStyles、主题fontScheme等回退。
    其它属性（字号/颜色/粗斜体/下划线/删除线）维持原有 run/para 级别解析。
    """
    props: Dict[str, Any] = {}
    try:
        rf = run.font
        pf = getattr(para, 'font', None)
        # 字体名解析仅限：run.rPr / run.font / para.font / para.pPr
        name = None
        name_src = None
        reason = None
        try:
            rPr = run._r.rPr if hasattr(run, '_r') else None
            if rPr is not None and getattr(rPr, 'rFonts', None) is not None:
                rfonts = rPr.rFonts
                # 依次尝试 eastAsia、ascii、hAnsi、cs、latin.typeface
                name = (
                    getattr(rfonts, 'eastAsia', None)
                    or getattr(rfonts, 'ascii', None)
                    or getattr(rfonts, 'hAnsi', None)
                    or getattr(rfonts, 'cs', None)
                )
                if name is None and getattr(rPr, 'latin', None) is not None:
                    name = getattr(rPr.latin, 'typeface', None)
                if name is not None:
                    name_src = 'run.rPr.rFonts/latin'
        except Exception:
            pass
        if name is None and getattr(rf, 'name', None) is not None:
            name = rf.name
            name_src = 'run.font.name'
        if name is None and pf is not None and getattr(pf, 'name', None) is not None:
            name = pf.name
            name_src = 'para.font.name'
        if name is None:
            val = _get_para_rfonts(para)
            if val is not None:
                name = val
                name_src = 'para.pPr.rPr.rFonts/latin'
        if name is None:
            reason = 'run/paragraph 层均未提供字体名'
        # 规范化字体名称（去除前后空白）；如果所有方法都无法获取字体名，则设为"未知"
        if isinstance(name, str):
            name = name.strip()
            if name.startswith('+'):
                # 主题占位符不再解析
                name = "未知"
                name_src = name_src + ' -> 未知' if name_src else '未知'
                reason = '主题占位符（+ 前缀）不做解析'
        elif name is None:
            # 所有方法都无法获取字体名，设为"未知"
            name = "未知"
            name_src = '未知'
            if reason is None:
                reason = '未从 run/para 获取到字体名'
        # 字体族合并（规范别名/派生名）
        merged = _merge_font_family_alias(name)
        props["字体类型"] = merged

        # 可选：未知时简单提示（保留最小化日志）
        try:
            if merged == "未知":
                sid = str(getattr(host_shape, "shape_id", ""))
                snippet = ''
                try:
                    snippet = (run.text or '')[:30]
                except Exception:
                    snippet = ''
                print(f"[字体类型=未知] shape_id={sid} 源={name_src} 原因={reason or '无'} 文本片段='{snippet}'")
        except Exception:
            pass
        # 字号
        size = getattr(rf, 'size', None) or (getattr(pf, 'size', None) if pf is not None else None)
        props["字号"] = float(size.pt) if size is not None else 18.0
        # 颜色
        color = getattr(rf, 'color', None)
        hex_color = _rgb_to_hex(color) if color is not None else None
        if hex_color is None and pf is not None:
            pcolor = getattr(pf, 'color', None)
            hex_color = _rgb_to_hex(pcolor) if pcolor is not None else None
        if hex_color is None:
            hex_color = "#000000"
        props["字体颜色"] = hex_color
        # 样式
        def _bool_or_default(val, fallback):
            return bool(val) if val is not None else fallback
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
            "是否粗体": _bool_or_default(bold, False),
            "是否斜体": _bool_or_default(italic, False),
            "是否下划线": _bool_or_default(underline, False),
            "是否带删除线": _bool_or_default(strike, False),
        })
    except Exception:
        pass
    return props


def _get_text_block_info(shape, shape_index: int) -> Dict[str, Any]:
    text_info = {}
    try:
        # 检查是否为组合元素
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # 递归处理组合元素中的子形状
            group_text_info = {}
            for i, sub_shape in enumerate(shape.shapes):
                # 使用数字索引，避免字符串拼接问题
                sub_text_info = _get_text_block_info(sub_shape, shape_index * 100 + i)
                if sub_text_info:
                    group_text_info.update(sub_text_info)
            return group_text_info
        
        # 处理普通文本形状
        if shape.has_text_frame and shape.text_frame:
            text_block_position = _get_shape_position(shape)
            is_title_placeholder = _is_title_placeholder(shape)
            
            # 构建文本块数据（改：输出“段落属性”，不再生成“拼接字符”）
            text_block_data = {
                "文本块位置": text_block_position,
                "图层编号": shape_index,
                "是否是标题占位符": is_title_placeholder,
                "文本块索引": str(getattr(shape, "shape_id", "")),
                "段落属性": []
            }
            
            # 获取文本框架中的段落和运行
            text_frame = shape.text_frame
            for para_index, paragraph in enumerate(text_frame.paragraphs):
                # 处理段落中的每个运行（run）
                for run_index, run in enumerate(paragraph.runs):
                    # 获取运行的字体属性
                    font_props = _resolve_run_font_props(run, paragraph, is_title_placeholder, shape)
                    
                    # 构建 run 级属性对象（段落属性项）
                    char_attr = {
                        "段落编号": para_index,
                        "字体类型": font_props.get("字体类型", "未知"),
                        "字号": font_props.get("字号", 18.0),
                        # 颜色改为中文常见色名
                        "字体颜色": _hex_to_cn_color_name(font_props.get("字体颜色", "#000000")),
                        "是否粗体": font_props.get("是否粗体", False),
                        # "是否斜体": font_props.get("是否斜体", False),
                        # "是否下划线": font_props.get("是否下划线", False),
                        # "是否带删除线": font_props.get("是否带删除线", False),
                        "段落内容": run.text
                    }

                    # JSON层面合并：若与前一条在样式上完全一致（仅段落内容不同），则合并段落内容
                    if text_block_data["段落属性"]:
                        last = text_block_data["段落属性"][-1]
                        same_style = all(last.get(k) == char_attr.get(k) for k in ATTR_COMPARE_KEYS)
                        same_para = last.get("段落编号") == char_attr.get("段落编号")
                        if same_style and same_para:
                            last["段落内容"] = f"{last.get('段落内容','')}{char_attr.get('段落内容','')}"
                        else:
                            text_block_data["段落属性"].append(char_attr)
                    else:
                        text_block_data["段落属性"].append(char_attr)

            # 检查是否有有效的段落内容
            has_content = False
            for char_attr in text_block_data["段落属性"]:
                if char_attr.get("段落内容", "").strip():
                    has_content = True
                    break
            
            # 只有当有内容时才输出文本块
            if has_content:
                text_key = f"文本块{shape_index + 1}"
                sid = str(getattr(shape, "shape_id", ""))
                text_payload = {
                    "文本块位置": text_block_position,
                    "图层编号": shape_index,
                    "是否是标题占位符": is_title_placeholder,
                    "文本块索引": sid,
                    "段落属性": text_block_data["段落属性"]
                }
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
            # 提取 shape_id 作为图片索引（便于与原始PPT形状对应）
            sid = str(getattr(shape, "shape_id", ""))
            image_info = {
                f"图片{shape_index + 1}": {
                    "图片位置": position,
                    "图片类型": image_type,
                    "图片大小": f"{image_size} bytes",
                    "图层位置": shape_index,
                    "图片索引": sid
                }
            }
    except Exception as e:
        print(f"提取图片信息失败: {e}")
    return image_info


def parse_pptx(path: str, include_images: bool = False) -> Dict[str, Any]:
    try:
        prs = Presentation(path)
        total_slides = len(prs.slides)
        
        # 按照新结构组织数据
        result = {
            "页数": total_slides,
            "contents": []
        }
        
        for slide_index, slide in enumerate(prs.slides):
            page_data = {
                "页码": slide_index + 1,
                "文本块数量": 0,
                "文本块": [],
                "图片数量": 0,
                "图片": []
            }
            
            # 处理文本块（包括组合元素中的文本）
            text_blocks: List[Dict[str, Any]] = []
            for shape_index, shape in enumerate(slide.shapes):
                text_info = _get_text_block_info(shape, shape_index)
                if text_info:
                    # 提取文本块内容到数组
                    for key, payload in text_info.items():
                        if key.startswith("文本块"):
                            text_blocks.append(payload)
            
            page_data["文本块数量"] = len(text_blocks)
            page_data["文本块"] = text_blocks
            
            # 处理图片
            if include_images:
                images: List[Dict[str, Any]] = []
                for shape_index, shape in enumerate(slide.shapes):
                    image_info = _get_image_info(shape, shape_index)
                    if image_info:
                        # 提取图片内容到数组
                        for key, payload in image_info.items():
                            if key.startswith("图片"):
                                images.append(payload)
                
                page_data["图片数量"] = len(images)
                page_data["图片"] = images
            
            result["contents"].append(page_data)
        
        return result
    except Exception as e:
        print(f"解析PPTX文件失败: {e}")
        return {"页数": 0, "contents": []}


def save_to_json(data: List[Dict[str, Any]], output_path: str):
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"数据已保存到: {output_path}")
    except Exception as e:
        print(f"保存JSON文件失败: {e}")


if __name__ == "__main__":
    # 命令行参数：--include-images 控制是否输出图片信息（默认不输出）
    import argparse
    parser = argparse.ArgumentParser(description="PPTX解析器")
    parser.add_argument("pptx", nargs='?', default="example2.pptx", help="PPTX 文件路径")
    parser.add_argument("--include-images", action="store_true", help="是否输出图片信息，默认否")
    args = parser.parse_args()

    pptx_path = args.pptx
    print("🧪 测试PPTX解析...")
    # 依据新增参数 include_images 控制图片信息输出
    result = parse_pptx(pptx_path, include_images=args.include_images)
    if result and "contents" in result:
        print(f"✅ 成功解析，共 {result['页数']} 页")
        save_to_json(result, "parsing_result.json")
        
        # # 显示前几页的关键信息
        # for i, page in enumerate(result['contents'][:5]):
        #     print(f"\n第 {page['页码']} 页:")
            
        #     # 显示前几个文本块的关键信息
        #     for j, text_block in enumerate(page.get('文本块', [])):  # 只显示前3个文本块
        #         print(text_block)
    else:
        print("❌ 解析失败")

