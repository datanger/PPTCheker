"""
PPTX 解析器 - 提取PPT文件的详细信息

输出格式：
- 每页为一个对象映射：{"文本块N": {...}, "图片M": {...}}
- 文本块：文本块位置、图层编号、是否是标题占位符、字符属性数组、拼接字符
- 图片：图片位置、类型、大小、图层位置

注意：某些字段可能无法直接获取，已删除并说明原因
"""

import json
from typing import List, Dict, Any, Optional, Tuple
from pptx import Presentation
from pptx.dml.color import RGBColor, MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE

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
# 已移除：DEBUG 输出、演示级 defaultTextStyle 缓存与加载、外部文件修补函数


def patch_theme_eastasia_fonts(*args, **kwargs) -> bool:
    """已删除冗余外部修补逻辑，占位返回False。"""
    return False


def patch_presentation_defaulttextstyle_ea(*args, **kwargs) -> bool:
    """已删除冗余外部修补逻辑，占位返回False。"""
    return False


def patch_all_lststyle_eastasia(*args, **kwargs) -> bool:
    """已删除冗余外部修补逻辑，占位返回False。"""
    return False


def patch_master_title_eastasia(*args, **kwargs) -> bool:
    """已删除冗余外部修补逻辑，占位返回False。"""
    return False



# 主题占位符字体名默认映射（当主题未给出 eastAsia/latin 实际字体时的兜底）
# 说明：
# - "+mn-ea"/"+mj-ea" 视为东亚字体，默认映射到“微软雅黑”
# - "+mn-lt"/"+mj-lt" 视为拉丁字体，默认映射到“Calibri”
THEME_PLACEHOLDER_DEFAULT_MAP = {
    "+mn-ea": "微软雅黑",
    "+mj-ea": "微软雅黑",
    "+mn-lt": "Calibri",
    "+mj-lt": "Calibri",
}


def _map_theme_placeholder_to_font(name: str, theme_fonts: Dict[str, Optional[str]]) -> Optional[str]:
    # 仅当主题明确提供对应映射时返回，否则None（不做兜底猜测）
    if not isinstance(name, str) or not name.startswith('+'):
        return None
    low = name.lower()
    if 'mj' in low and 'ea' in low:
        return theme_fonts.get('major_eastAsia')
    if 'mn' in low and 'ea' in low:
        return theme_fonts.get('minor_eastAsia')
    if 'mj' in low and ('lt' in low or 'latin' in low):
        return theme_fonts.get('major_latin')
    if 'mn' in low and ('lt' in low or 'latin' in low):
        return theme_fonts.get('minor_latin')
    return None


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


def _get_shape_position(shape) -> Dict[str, str]:
    """返回形状位置，单位为毫米（mm），并在值中附带单位字符串。
    PowerPoint内部单位为EMU，换算：1 mm = 36000 EMU。
    """
    try:
        def emu_to_mm_str(emu_val) -> str:
            try:
                mm = float(emu_val) / 36000.0
                return f"{mm:.2f} mm"
            except Exception:
                return "0.00 mm"
        return {
            "left": emu_to_mm_str(shape.left),
            "top": emu_to_mm_str(shape.top),
            "width": emu_to_mm_str(shape.width),
            "height": emu_to_mm_str(shape.height)
        }
    except Exception:
        return {"left": "0.00 mm", "top": "0.00 mm", "width": "0.00 mm", "height": "0.00 mm"}


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


def _extract_font_props_from_font_obj(font_obj) -> Dict[str, Any]:
    props: Dict[str, Any] = {}
    if font_obj is None:
        return props
    try:
        size = getattr(font_obj, 'size', None)
        color = getattr(font_obj, 'color', None)
        props.update({
            "字体类型": getattr(font_obj, 'name', None),
            "字号": float(size.pt) if size is not None else None,
            "字体颜色": _rgb_to_hex(color) if color is not None else None,
            "是否粗体": getattr(font_obj, 'bold', None),
            "是否斜体": getattr(font_obj, 'italic', None),
            "是否下划线": getattr(font_obj, 'underline', None),
            "是否带删除线": getattr(font_obj, 'strike', None),
        })
    except Exception:
        pass
    return props


def _inherit_placeholder_defaults(shape) -> Dict[str, Any]:
    """从版式/母版占位符继承字体默认值。"""
    defaults: Dict[str, Any] = {}
    try:
        if not getattr(shape, 'is_placeholder', False):
            return defaults
        phf = shape.placeholder_format
        ph_idx = getattr(phf, 'idx', None)
        slide_layout = getattr(shape.part, 'slide_layout', None)
        # 1) 版式占位符
        try:
            if slide_layout is not None:
                for p in slide_layout.placeholders:
                    try:
                        if getattr(p.placeholder_format, 'idx', None) == ph_idx:
                            # 取版式占位符的段落字体
                            if hasattr(p, 'text_frame') and p.text_frame and p.text_frame.paragraphs:
                                df = _extract_font_props_from_font_obj(p.text_frame.paragraphs[0].font)
                                defaults.update({k: v for k, v in df.items() if v is not None})
                            break
                    except Exception:
                        continue
        except Exception:
            pass
        # 2) 母版占位符
        try:
            master = getattr(slide_layout, 'slide_master', None) if slide_layout is not None else None
            if master is not None:
                for p in master.placeholders:
                    try:
                        if getattr(p.placeholder_format, 'idx', None) == ph_idx:
                            if hasattr(p, 'text_frame') and p.text_frame and p.text_frame.paragraphs:
                                df = _extract_font_props_from_font_obj(p.text_frame.paragraphs[0].font)
                                defaults.update({k: v for k, v in df.items() if v is not None})
                            break
                    except Exception:
                        continue
        except Exception:
            pass
    except Exception:
        pass
    return defaults


def _get_theme_major_minor_fonts(shape) -> Dict[str, Optional[str]]:
    """尝试从主题(fontScheme)获取major/minor字体（latin/eastAsia）。"""
    result = {"major_latin": None, "major_eastAsia": None, "minor_latin": None, "minor_eastAsia": None}
    try:
        slide_layout = getattr(shape.part, 'slide_layout', None)
        slide_master = getattr(slide_layout, 'slide_master', None) if slide_layout is not None else None
        theme_part = getattr(slide_master.part, 'theme_part', None) if slide_master is not None else None
        theme = getattr(theme_part, 'theme', None) if theme_part is not None else None
        font_scheme = getattr(theme, 'fontScheme', None) if theme is not None else None
        if font_scheme is None:
            # python-pptx对象模型可能不同，尝试从element层访问
            try:
                theme_el = theme_part._element  # lxml element
                # 查找a:fontScheme 下的a:majorFont/a:minorFont
                for child in theme_el.iter():
                    tag = child.tag.lower()
                    if tag.endswith('majorfont'):
                        for f in child:
                            ftag = f.tag.lower()
                            if ftag.endswith('latin') and 'typeface' in f.attrib:
                                result['major_latin'] = f.attrib.get('typeface')
                            if ftag.endswith('ea') and 'typeface' in f.attrib:  # eastAsia
                                result['major_eastAsia'] = f.attrib.get('typeface')
                    if tag.endswith('minorfont'):
                        for f in child:
                            ftag = f.tag.lower()
                            if ftag.endswith('latin') and 'typeface' in f.attrib:
                                result['minor_latin'] = f.attrib.get('typeface')
                            if ftag.endswith('ea') and 'typeface' in f.attrib:
                                result['minor_eastAsia'] = f.attrib.get('typeface')
            except Exception:
                pass
        else:
            try:
                major = font_scheme.majorFont
                minor = font_scheme.minorFont
                # 这些属性名依赖python-pptx实现，做异常保护
                result['major_latin'] = getattr(major, 'latin', None) and getattr(major.latin, 'typeface', None)
                result['minor_latin'] = getattr(minor, 'latin', None) and getattr(minor.latin, 'typeface', None)
            except Exception:
                pass
    except Exception:
        pass
    return result


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


def _get_shape_lststyle_font(shape, para_level: int) -> Optional[str]:
    """从 shape.text_frame._txBody.lstStyle 的 lvlXpPr.defRPr.rFonts/latin 取字体。"""
    try:
        tf = getattr(shape, 'text_frame', None)
        txBody = getattr(tf, '_txBody', None) if tf is not None else None
        lstStyle = getattr(txBody, 'lstStyle', None) if txBody is not None else None
        if lstStyle is None:
            return None
        lvl_idx = max(1, (para_level or 0) + 1)
        # 属性名如 lvl1pPr, lvl2pPr
        lvl_attr = f'lvl{lvl_idx}pPr'
        lvl = getattr(lstStyle, lvl_attr, None)
        defrpr = None
        if lvl is not None:
            defrpr = getattr(lvl, 'defRPr', None) or getattr(lvl, 'rPr', None)
        if defrpr is None:
            # 直接在 lstStyle 下找 defRPr
            defrpr = getattr(lstStyle, 'defRPr', None) or getattr(lstStyle, 'rPr', None)
        if defrpr is None:
            return None
        rfonts = getattr(defrpr, 'rFonts', None)
        if rfonts is not None:
            name = (
                getattr(rfonts, 'eastAsia', None)
                or getattr(rfonts, 'ascii', None)
                or getattr(rfonts, 'hAnsi', None)
                or getattr(rfonts, 'cs', None)
            )
            if name:
                return name
        latin = getattr(defrpr, 'latin', None)
        if latin is not None and getattr(latin, 'typeface', None):
            return latin.typeface
    except Exception:
        pass
    return None


# 注：移除启发式字体猜测，避免误判。

def _get_master_textstyle_font(shape, para_level: int) -> Optional[str]:
    """从母版 textStyles 中按段落层级提取缺省字体(typeface)。
    优先顺序：bodyStyle → headingStyle → titleStyle；层级映射 lvl{n}pPr，n=para_level+1。
    返回 eastAsia/ascii/hAnsi/cs/latin 中优先可用的 typeface。
    """
    try:
        slide_layout = getattr(shape.part, 'slide_layout', None)
        slide_master = getattr(slide_layout, 'slide_master', None) if slide_layout is not None else None
        if slide_master is None:
            return None
        root = getattr(slide_master, '_element', None)
        if root is None:
            return None
        # 查找 a:txStyles
        tx_styles = None
        for el in root.iter():
            tag = el.tag.lower()
            if tag.endswith('txstyles'):
                tx_styles = el
                break
        if tx_styles is None:
            return None
        # 目标层级标签名，如 lvl1ppr, lvl2ppr ...
        lvl_tag = f'lvl{max(1, (para_level or 0) + 1)}ppr'
        # 在 bodyStyle/headingStyle/titleStyle 依序查找
        sections = []
        for child in tx_styles:
            tag = child.tag.lower()
            if tag.endswith('bodystyle') or tag.endswith('headingstyle') or tag.endswith('titlestyle'):
                sections.append(child)
        # 若没有显式 section，则允许直接在 txStyles 下查 lvlXpPr/defRPr
        if not sections:
            sections = [tx_styles]

        for sec_el in sections:
            # 寻找层级 pPr（可能直接在 txStyles 下，或在 section 下）
            lvl_el = None
            for child in sec_el:
                if child.tag.lower().endswith(lvl_tag):
                    lvl_el = child
                    break
            # 若未找到层级，尝试默认 defRPr
            defrpr = None
            if lvl_el is not None:
                for c in lvl_el:
                    if c.tag.lower().endswith('defrpr') or c.tag.lower().endswith('rpr'):
                        defrpr = c
                        break
            if defrpr is None:
                for child in sec_el:
                    if child.tag.lower().endswith('defrpr') or child.tag.lower().endswith('rpr'):
                        defrpr = child
                        break
            if defrpr is None:
                continue
            # 读取 rFonts/latin@typeface
            east_asia = ascii_v = hansi = cs = latin = None
            for c in defrpr:
                t = c.tag.lower()
                if t.endswith('rfonts'):
                    east_asia = c.attrib.get('eastasia') or c.attrib.get('ea')
                    ascii_v = c.attrib.get('ascii')
                    hansi = c.attrib.get('hansi')
                    cs = c.attrib.get('cs')
                if t.endswith('latin') and 'typeface' in c.attrib:
                    latin = c.attrib.get('typeface')
            name = east_asia or ascii_v or hansi or cs or latin
            if name:
                return name
        return None
    except Exception:
        return None


def _get_master_title_font(shape, para_level: int) -> Optional[str]:
    """专取母版 titleStyle 的默认字体。"""
    try:
        slide_layout = getattr(shape.part, 'slide_layout', None)
        slide_master = getattr(slide_layout, 'slide_master', None) if slide_layout is not None else None
        if slide_master is None:
            return None
        root = getattr(slide_master, '_element', None)
        if root is None:
            return None
        tx_styles = None
        for el in root.iter():
            if el.tag.lower().endswith('txstyles'):
                tx_styles = el
                break
        if tx_styles is None:
            return None
        lvl_tag = f'lvl{max(1, (para_level or 0) + 1)}ppr'
        title_sec = None
        for child in tx_styles:
            if child.tag.lower().endswith('titlestyle'):
                title_sec = child
                break
        if title_sec is None:
            return None
        lvl_el = None
        for c in title_sec:
            if c.tag.lower().endswith(lvl_tag):
                lvl_el = c
                break
        defrpr = None
        if lvl_el is not None:
            for c in lvl_el:
                t = c.tag.lower()
                if t.endswith('defrpr') or t.endswith('rpr'):
                    defrpr = c
                    break
        if defrpr is None:
            for c in title_sec:
                t = c.tag.lower()
                if t.endswith('defrpr') or t.endswith('rpr'):
                    defrpr = c
                    break
        if defrpr is None:
            return None
        east_asia = ascii_v = hansi = cs = latin = None
        for c in defrpr:
            t = c.tag.lower()
            if t.endswith('rfonts'):
                east_asia = c.attrib.get('eastasia') or c.attrib.get('ea')
                ascii_v = c.attrib.get('ascii')
                hansi = c.attrib.get('hansi')
                cs = c.attrib.get('cs')
            if t.endswith('latin') and 'typeface' in c.attrib:
                latin = c.attrib.get('typeface')
        return east_asia or ascii_v or hansi or cs or latin
    except Exception:
        return None


def _resolve_run_font_props(run, para, is_title_placeholder: bool, host_shape) -> Dict[str, Any]:
    """解析run的有效字体属性：名称、字号、颜色、粗体、斜体、下划线、删除线。
    规则：优先run.font，其次para.font，再次版式/母版占位符默认，最后回退默认值。
    颜色支持主题色映射与亮度换算。缺省回退：字体名“默认”、字号18pt、颜色#000000、布尔样式False。
    """
    props: Dict[str, Any] = {}
    try:
        rf = run.font
        pf = getattr(para, 'font', None)
        placeholder_defaults = _inherit_placeholder_defaults(run._r.getparent().getparent()) if hasattr(run, '_r') else {}
        theme_fonts = _get_theme_major_minor_fonts(run._r.getparent().getparent()) if hasattr(run, '_r') else {}
        # 字体名解析优先级：
        # 1) 显式 rFonts（run级）
        # 2) run.font
        # 3) para.font
        # 4) 段落 pPr.rPr.rFonts / latin
        # 5) 占位符默认
        # 6) shape 的 lstStyle（txBody）
        # 7) 母版 textStyles
        # 8) 主题 major/minor
        # 9) 脚本启发式猜测（最后）
        name = None
        name_src = None
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
        if name is None and placeholder_defaults.get("字体类型") is not None:
            name = placeholder_defaults.get("字体类型")
            name_src = 'placeholder_defaults'
        if name is None:
            try:
                level = getattr(para, 'level', 0)
                val = _get_shape_lststyle_font(run._r.getparent().getparent(), level) if hasattr(run, '_r') else None
                if val is not None:
                    name = val
                    name_src = 'shape.lstStyle'
            except Exception:
                pass
        # 母版 textStyles 作为进一步回退
        if name is None:
            try:
                level = getattr(para, 'level', 0)
                val = _get_master_textstyle_font(run._r.getparent().getparent(), level) if hasattr(run, '_r') else None
                if val is not None:
                    name = val
                    name_src = 'master.txStyles'
            except Exception:
                pass
        # 文档级 defaultTextStyle 回退（presentation.xml）
        # 已移除 presentation.defaultTextStyle 回退
        # 主题字体方案作为最后回退（若仍为 None）
        if name is None:
            val = (
                theme_fonts.get('major_eastAsia')
                or theme_fonts.get('major_latin')
                or theme_fonts.get('minor_eastAsia')
                or theme_fonts.get('minor_latin')
            )
            if val is not None:
                name = val
                name_src = 'theme.fontScheme'
        # 若是标题占位符，优先尝试母版 titleStyle（只做确定性解析）
        if (name is None or (isinstance(name, str) and name.startswith('+'))) and is_title_placeholder:
            try:
                level = getattr(para, 'level', 0)
                val = _get_master_title_font(host_shape, level)
                if val:
                    name = val
                    name_src = (name_src + ' -> master.titleStyle') if name_src else 'master.titleStyle'
            except Exception:
                pass
        # 规范化字体名称（去除前后空白）；可能为 None（用于后续继承）
        if isinstance(name, str):
            name = name.strip()
            if name.startswith('+'):
                # 将主题占位符映射到具体字体（主题优先，内置兜底）；若仍无法映射，则保留占位符原值，供上层归一化
                mapped = _map_theme_placeholder_to_font(name, theme_fonts)
                if mapped is not None:
                    name = mapped
                    name_src = name_src + ' -> theme_placeholder_map' if name_src else 'theme_placeholder_map'
        props["字体类型"] = name
        # 字号
        size = getattr(rf, 'size', None) or (getattr(pf, 'size', None) if pf is not None else None)
        if size is None and placeholder_defaults.get("字号") is not None:
            props["字号"] = float(placeholder_defaults.get("字号"))
        else:
            props["字号"] = float(size.pt) if size is not None else 18.0
        # 颜色
        color = getattr(rf, 'color', None)
        hex_color = _rgb_to_hex(color) if color is not None else None
        if hex_color is None and pf is not None:
            pcolor = getattr(pf, 'color', None)
            hex_color = _rgb_to_hex(pcolor) if pcolor is not None else None
        if hex_color is None and placeholder_defaults.get("字体颜色") is not None:
            hex_color = placeholder_defaults.get("字体颜色")
        if hex_color is None:
            hex_color = "#000000"
        props["字体颜色"] = hex_color
        # 样式
        def _bool_or_default(val, fallback):
            return bool(val) if val is not None else fallback
        bold = getattr(rf, 'bold', None)
        if bold is None and pf is not None:
            bold = getattr(pf, 'bold', None)
        if bold is None:
            bold = placeholder_defaults.get("是否粗体")
        italic = getattr(rf, 'italic', None)
        if italic is None and pf is not None:
            italic = getattr(pf, 'italic', None)
        if italic is None:
            italic = placeholder_defaults.get("是否斜体")
        underline = getattr(rf, 'underline', None)
        if underline is None and pf is not None:
            underline = getattr(pf, 'underline', None)
        if underline is None:
            underline = placeholder_defaults.get("是否下划线")
        strike = getattr(rf, 'strike', None)
        if strike is None and pf is not None:
            strike = getattr(pf, 'strike', None)
        if strike is None:
            strike = placeholder_defaults.get("是否带删除线")
        props.update({
            "是否粗体": _bool_or_default(bold, False),
            "是否斜体": _bool_or_default(italic, False),
            "是否下划线": _bool_or_default(underline, False),
            "是否带删除线": _bool_or_default(strike, False),
        })
        # 已移除字体调试输出
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
                    eff = _resolve_run_font_props(run, para, is_title_placeholder, shape)
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


def parse_pptx(path: str, include_images: bool = False) -> List[Dict[str, Any]]:
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
                if include_images:
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
    # 命令行参数：--include-images 控制是否输出图片信息（默认不输出）
    import argparse
    parser = argparse.ArgumentParser(description="PPTX解析器")
    parser.add_argument("pptx", nargs='?', default="example1.pptx", help="PPTX 文件路径")
    parser.add_argument("--include-images", action="store_true", help="是否输出图片信息，默认否")
    args = parser.parse_args()

    pptx_path = args.pptx
    print("🧪 测试PPTX解析...")
    # 依据新增参数 include_images 控制图片信息输出
    result = parse_pptx(pptx_path, include_images=args.include_images)
    if result:
        print(f"✅ 成功解析，共 {len(result)} 页")
        save_to_json(result, "parsing_result_new.json")
        first_page = result[0]
        print(f"\n📄 第一页包含 {len(first_page)} 个元素")
        for i in range(len(result[:5])):
            print(f"第 {i+1} 页:")
            for k, v in result[i].items():
                v.pop("字符属性")
                print(k, v)
                # break
                # if k.startswith("文本块"):
                #     s = v.get("拼接字符", "")
                #     print(s)    
                #     # break
    else:
        print("❌ 解析失败")

