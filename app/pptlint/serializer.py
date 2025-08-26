"""
基于元数据的字符属性串序列化器（差分变更模式）

规则：
- 初始输出一次全量7项属性，标记为【初始的字符所有属性：...】
- 后续仅当属性发生变化时输出“变更项”，标记为【字符属性变更：...】
- 变更标记中只包含发生变化的键，不重复未变更项
- 仍支持【换行】【缩进{N}】控制标记
- 颜色标准化为十六进制；初始属性采用整块最常见非空属性并补齐
"""

from typing import List, Dict, Any, Tuple, Optional
import json
import os
import re
from collections import Counter

ATTR_KEYS = [
    "字体类型",
    "字号",
    "字体颜色",
    "是否粗体",
    "是否斜体",
    "是否下划线",
    "是否带删除线",
]

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
    if value is True or value == 1:
        return "是"
    if value is False or value == 0:
        return "否"
    return "否"  # 三态缺省按否处理


def _to_hex_color(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, str) and re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        v = value if value.startswith('#') else f"#{value}"
        return v.upper()
    if isinstance(value, str) and value.lower().startswith("rgb"):
        nums = re.findall(r"\d+", value)
        if len(nums) >= 3:
            r, g, b = (int(nums[0]), int(nums[1]), int(nums[2]))
            return f"#{r:02X}{g:02X}{b:02X}"
    if isinstance(value, str) and value in NAMED_COLOR_TO_HEX:
        return NAMED_COLOR_TO_HEX[value]
    if isinstance(value, (tuple, list)) and len(value) >= 3:
        try:
            r, g, b = int(value[0]), int(value[1]), int(value[2])
            return f"#{r:02X}{g:02X}{b:02X}"
        except Exception:
            pass
    if isinstance(value, str):
        for name, hexv in NAMED_COLOR_TO_HEX.items():
            if name in value:
                return hexv
    return None


def _normalize_value(key: str, value: Any) -> Any:
    if key in ("是否粗体", "是否斜体", "是否下划线", "是否带删除线"):
        return _format_bool(value)
    if key == "字体颜色":
        hexv = _to_hex_color(value)
        return hexv if hexv is not None else "#000000"
    if key == "字体类型":
        # 将字体类型归一为有限集合：微软雅黑、宋体、Meiryo UI、楷体、timenew roman；其余为“其他”
        if not isinstance(value, str) or not value.strip():
            return "其他"
        raw = value.strip().lower().replace(" ", "")
        # 常见别名/英文字体族名归一
        if raw in {"微软雅黑", "microsoftyahei", "msyahei", "yahei"}:
            return "微软雅黑"
        if raw in {"宋体", "simsun"}:
            return "宋体"
        if raw in {"楷体", "kaiti", "kaitisc", "stkaiti"}:
            return "楷体"
        # Meiryo UI 系列别名：含 meiryo/meiyou/拼写变体，统一归为 Meiryo UI
        if raw in {"meiryoui", "meiryoui", "meiryo", "meiyou"} or value.strip() in {"Meiryo UI", "Meiryo"}:
            return "Meiryo UI"
        if raw in {"timesnewroman", "timenewroman", "timesnewromanpsmt"} or value.strip() in {"Times New Roman", "TimeNew Roman", "timenew roman"}:
            return "timenew roman"
        # 若为主题占位符（+mn-ea/+mj-lt 等），归为“其他”（保持不猜测）
        if raw.startswith('+'):
            return "其他"
        return "其他"
    return value


def _attrs_from_char(char_info: Dict[str, Any], prev: Dict[str, Any]) -> Dict[str, Any]:
    snapshot: Dict[str, Any] = {}
    for k in ATTR_KEYS:
        if k in char_info:
            snapshot[k] = char_info[k]
        else:
            if k in prev:
                snapshot[k] = prev[k]
            else:
                snapshot[k] = None
    return snapshot


def _make_initial_marker(attrs: Dict[str, Any], label: str) -> str:
    parts = []
    for k in ATTR_KEYS:
        v = _normalize_value(k, attrs.get(k))
        parts.append(f"{k}{{{v}}}")
    return f"【{label}：" + "、".join(parts) + "】"


def _make_full_attrs_marker(attrs: Dict[str, Any]) -> str:
    parts = []
    for k in ATTR_KEYS:
        v = _normalize_value(k, attrs.get(k))
        parts.append(f"{k}{{{v}}}")
    return "【字符属性：" + "、".join(parts) + "】"


def _make_changed_attrs_marker(prev_attrs: Dict[str, Any], curr_attrs: Dict[str, Any]) -> Optional[str]:
    """仅格式化变更的属性键值；无变更则返回 None。"""
    changed_parts = []
    for k in ATTR_KEYS:
        pv = prev_attrs.get(k)
        cv = curr_attrs.get(k)
        if pv != cv:
            v = _normalize_value(k, cv)
            changed_parts.append(f"{k}{{{v}}}")
    if not changed_parts:
        return None
    return "【字符属性变更：" + "、".join(changed_parts) + "】"


def _most_frequent_non_unknown(values: List[Any]) -> Optional[Any]:
    filtered = [v for v in values if v is not None]
    if not filtered:
        return None
    cnt = Counter(filtered)
    best, _ = cnt.most_common(1)[0]
    return best


def _build_initial_attrs(chars: List[Dict[str, Any]]) -> Dict[str, Any]:
    """以首个可见字符作为初始属性基线。"""
    if not chars:
        return {k: None for k in ATTR_KEYS}
    first = next((c for c in chars if c.get("字符内容") != "\n"), None)
    if first is None:
        return {k: None for k in ATTR_KEYS}
    return {k: first.get(k) for k in ATTR_KEYS}


def serialize_text_block_to_diff_string(text_block: Dict[str, Any], initial_label: str = "初始的字符所有属性") -> str:
    text_info = next(iter(text_block.values()))
    chars: List[Dict[str, Any]] = text_info.get("字符属性", [])
    if not chars:
        return _make_initial_marker({k: None for k in ATTR_KEYS}, initial_label)

    output_parts: List[str] = []
    # 计算初始属性，并标准化用于后续比较
    initial_attrs_raw = _build_initial_attrs(chars)
    initial_attrs = {k: _normalize_value(k, initial_attrs_raw.get(k)) for k in ATTR_KEYS}
    prev_attrs: Dict[str, Any] = {}
    at_baseline: bool = True  # 是否处于“与初始属性一致”的状态
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
            # 基于与初始属性的比较，更新基线状态
            at_baseline = all(curr_attrs.get(k) == initial_attrs.get(k) for k in ATTR_KEYS)
            # 若一开始就偏离初始属性，则输出一次相对初始的变更说明
            if not at_baseline:
                if buf:
                    output_parts.append("".join(buf))
                    buf = []
                marker = _make_changed_attrs_marker(initial_attrs, curr_attrs)
                if marker:
                    output_parts.append(marker)
            buf.append(char_text)
        else:
            # 仅当“相对初始属性发生偏离，且此前处于基线状态”时输出变更说明
            is_baseline_now = all(curr_attrs.get(k) == initial_attrs.get(k) for k in ATTR_KEYS)
            if at_baseline and (not is_baseline_now):
                if buf:
                    output_parts.append("".join(buf))
                    buf = []
                marker = _make_changed_attrs_marker(initial_attrs, curr_attrs)
                if marker:
                    output_parts.append(marker)
            # 更新基线状态与上一属性快照
            at_baseline = is_baseline_now
            prev_attrs = curr_attrs
            buf.append(char_text)
        i += 1

    if buf:
        output_parts.append("".join(buf))

    return "".join(output_parts)


def serialize_metadata_to_diff_strings(metadata: List[List[Dict[str, Any]]], initial_label: str = "初始的字符所有属性") -> List[List[str]]:
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
