"""
基于元数据的字符属性差分串序列化器

对应计划部分：
- 将“元数据”进行整合，按字符属性变更输出单字符串表达，减少tokens
- 仅在属性变化处输出【字符属性变更：...】标记
- 识别换行与段首缩进，输出【换行】【缩进{N}】
- 颜色标准化为十六进制；初始属性中的未知项尽量用后续字符的可获取值补齐

使用方法：
- 从 parsing_result_new.json 加载元数据
- 调用 serialize_metadata_to_diff_strings，得到每页每文本块的差分字符串
"""

from typing import List, Dict, Any, Tuple, Optional
import json
import os
import re
from collections import Counter

# 需要聚合的字符属性键（与元数据一致）
ATTR_KEYS = [
    "字体类型",
    "字号",
    "字体颜色",
    "是否粗体",
    "是否斜体",
    "是否下划线",
    "是否带删除线",
]

# 常见中文颜色名映射
NAMED_COLOR_TO_HEX = {
    "黑": "#000000", "黑色": "#000000",
    "白": "#FFFFFF", "白色": "#FFFFFF",
    "红": "#FF0000", "红色": "#FF0000",
    "绿": "#00FF00", "绿色": "#00FF00",
    "蓝": "#0000FF", "蓝色": "#0000FF",
    "黄": "#FFFF00", "黄色": "#FFFF00",
    "橙": "#FFA500", "橙色": "#FFA500",
    "紫": "#800080", "紫色": "#800080",
    "灰": "#808080", "灰色": "#808080",
}


def _format_bool(value: Any) -> str:
    """将布尔/数值转中文“是/否/未知”。"""
    if value is True or value == 1:
        return "是"
    if value is False or value == 0:
        return "否"
    return "未知"


def _to_hex_color(value: Any) -> Optional[str]:
    """将颜色值标准化为 #RRGGBB。支持：hex、rgb()/rgba() 字符串、中文名、(r,g,b)/[r,g,b]。"""
    if value is None:
        return None
    # 直接是hex
    if isinstance(value, str) and re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        v = value if value.startswith('#') else f"#{value}"
        return v.upper()
    # rgb()/rgba()
    if isinstance(value, str) and value.lower().startswith("rgb"):
        nums = re.findall(r"\d+", value)
        if len(nums) >= 3:
            r, g, b = (int(nums[0]), int(nums[1]), int(nums[2]))
            return f"#{r:02X}{g:02X}{b:02X}"
    # 中文名
    if isinstance(value, str) and value in NAMED_COLOR_TO_HEX:
        return NAMED_COLOR_TO_HEX[value]
    # 三元组/列表
    if isinstance(value, (tuple, list)) and len(value) >= 3:
        try:
            r, g, b = int(value[0]), int(value[1]), int(value[2])
            return f"#{r:02X}{g:02X}{b:02X}"
        except Exception:
            pass
    # 其它字符串尝试从中文名中提取
    if isinstance(value, str):
        for name, hexv in NAMED_COLOR_TO_HEX.items():
            if name in value:
                return hexv
    return None


def _normalize_value(key: str, value: Any) -> Any:
    """规范化属性值，保证输出稳定。"""
    if value is None:
        return "未知"
    if key in ("是否粗体", "是否斜体", "是否下划线", "是否带删除线"):
        return _format_bool(value)
    if key == "字体颜色":
        hexv = _to_hex_color(value)
        return hexv if hexv is not None else "未知"
    # 字号/字体类型按原样输出
    return value


def _diff_attrs(prev: Dict[str, Any], curr: Dict[str, Any]) -> Dict[str, Any]:
    """计算属性差异，仅输出发生变化的键。缺失视为沿用上一次。"""
    diff: Dict[str, Any] = {}
    for k in ATTR_KEYS:
        prev_v = prev.get(k, None)
        curr_v = curr.get(k, prev_v)
        if curr_v != prev_v:
            diff[k] = _normalize_value(k, curr_v)
    return diff


def _attrs_from_char(char_info: Dict[str, Any], prev: Dict[str, Any]) -> Dict[str, Any]:
    """从字符元数据提取属性快照，缺失字段继承 prev。"""
    snapshot: Dict[str, Any] = {}
    for k in ATTR_KEYS:
        if k in char_info:
            snapshot[k] = char_info[k]
        else:
            if k in prev:
                snapshot[k] = prev[k]
    return snapshot


def _make_change_marker(changes: Dict[str, Any]) -> str:
    """生成变更标记字符串，例如：【字符属性变更：字号{9}、是否下划线{是}】"""
    if not changes:
        return ""
    parts = []
    for k in ATTR_KEYS:
        if k in changes:
            parts.append(f"{k}{{{_normalize_value(k, changes[k])}}}")
    return "【字符属性变更：" + "、".join(parts) + "】"


def _make_initial_marker(attrs: Dict[str, Any], label: str = "初始的字符属性说明") -> str:
    """生成初始属性说明，包含全部七个属性。标签可配置。"""
    parts = []
    for k in ATTR_KEYS:
        v = _normalize_value(k, attrs.get(k))
        parts.append(f"{k}{{{v}}}")
    return f"【{label}：" + "、".join(parts) + "】"


def _most_frequent_non_unknown(values: List[Any], key: str) -> Optional[Any]:
    """返回列表中出现频率最高的非未知值（按规范化前的原值判断，再在返回前做规范化）。"""
    filtered = [v for v in values if v is not None]
    if not filtered:
        return None
    cnt = Counter(filtered)
    best, _ = cnt.most_common(1)[0]
    return best


def _build_initial_attrs(chars: List[Dict[str, Any]]) -> Dict[str, Any]:
    """以文本块内最常见非空属性作为初始值；若无则回退到首字符+向后补齐；颜色统一为hex。"""
    if not chars:
        return {k: "未知" for k in ATTR_KEYS}

    # 统计整块属性分布
    dist: Dict[str, List[Any]] = {k: [] for k in ATTR_KEYS}
    for ch in chars:
        if ch.get("字符内容") == "\n":
            continue
        for k in ATTR_KEYS:
            dist[k].append(ch.get(k))

    base: Dict[str, Any] = {}
    for k in ATTR_KEYS:
        best = _most_frequent_non_unknown(dist[k], k)
        base[k] = best

    # 若仍有缺失，使用原回退策略：首字符值并向后补齐
    if any(v is None for v in base.values()):
        first = next((c for c in chars if c.get("字符内容") != "\n"), None)
        if first is not None:
            for k in ATTR_KEYS:
                if base[k] is None:
                    base[k] = first.get(k, None)
        for k in ATTR_KEYS:
            if base[k] is None:
                for ch in chars:
                    if ch.get("字符内容") == "\n":
                        continue
                    if k in ch and ch[k] is not None:
                        base[k] = ch[k]
                        break

    # 规范化
    for k in ATTR_KEYS:
        base[k] = _normalize_value(k, base[k])
    return base


def serialize_text_block_to_diff_string(text_block: Dict[str, Any], initial_label: str = "初始的字符属性说明") -> str:
    """将单个文本块的字符数组序列化为差分字符串。
    - 以【初始的字符属性说明：...】作为起始标记（包含全部七个属性，未知项尽量补齐）
    - 后续只在属性变化处插入【字符属性变更：...】
    - 遇到换行符（\n）输出【换行】并统计后续连续空格数量输出【缩进{N}】
    - 文本连续段合并输出，避免逐字符冗余
    initial_label: 初始属性标签文本（例如“初始的字符所有属性”）
    """
    text_info = next(iter(text_block.values()))  # 取到 {"文本块X": {...}} 的内部字典
    chars: List[Dict[str, Any]] = text_info.get("字符属性", [])
    if not chars:
        base_unknown = {k: "未知" for k in ATTR_KEYS}
        return _make_initial_marker(base_unknown, initial_label)

    output_parts: List[str] = []
    initial_attrs = _build_initial_attrs(chars)
    prev_attrs: Dict[str, Any] = {}
    buf: List[str] = []

    i = 0
    while i < len(chars):
        ch = chars[i]
        char_text = ch.get("字符内容", "")
        if char_text == "\n":
            if buf:
                output_parts.append("".join(buf))
                buf = []
            output_parts.append("【换行】")
            j = i + 1
            indent_count = 0
            while j < len(chars) and chars[j].get("字符内容", "") == " ":
                indent_count += 1
                j += 1
            if indent_count > 0:
                output_parts.append(f"【缩进{{{indent_count}}}】")
            i = j
            continue

        curr_attrs_raw = _attrs_from_char(ch, prev_attrs)
        curr_attrs = {k: _normalize_value(k, curr_attrs_raw.get(k)) for k in ATTR_KEYS}

        if prev_attrs == {}:
            output_parts.append(_make_initial_marker(initial_attrs, initial_label))
            prev_attrs = curr_attrs
            buf.append(char_text)
        else:
            changes = _diff_attrs(prev_attrs, curr_attrs)
            if changes:
                if buf:
                    output_parts.append("".join(buf))
                    buf = []
                output_parts.append(_make_change_marker(changes))
                prev_attrs = curr_attrs
            buf.append(char_text)

        i += 1

    if buf:
        output_parts.append("".join(buf))

    return "".join(output_parts)


def serialize_metadata_to_diff_strings(metadata: List[List[Dict[str, Any]]], initial_label: str = "初始的字符属性说明") -> List[List[str]]:
    slides_strings: List[List[str]] = []
    for slide in metadata:
        slide_strings: List[str] = []
        for elem in slide:
            key = next(iter(elem.keys()))
            if key.startswith("文本块"):
                s = serialize_text_block_to_diff_string(elem, initial_label=initial_label)
                slide_strings.append(s)
        slides_strings.append(slide_strings)
    return slides_strings


def main():
    in_path = "parsing_result_new.json"
    out_path = "diff_strings.json"
    if not os.path.exists(in_path):
        print(f"未找到元数据文件: {in_path}")
        return
    with open(in_path, "r", encoding="utf-8") as f:
        metadata = json.load(f)
    diff = serialize_metadata_to_diff_strings(metadata, initial_label="初始的字符所有属性")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(diff, f, ensure_ascii=False, indent=2)
    print(f"已生成差分字符串文件: {out_path}")


if __name__ == "__main__":
    main()
