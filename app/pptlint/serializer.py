"""
基于元数据的字符属性串序列化器（差分变更模式）

规则：
- 初始输出一次全量7项属性，标记为【初始的字符所有属性：...】
- 后续仅当属性发生变化时输出“变更项”，标记为【字符属性变更：...】
- 变更标记中只包含发生变化的键，不重复未变更项
- 仍支持【换行】【缩进{N}】控制标记
- 颜色标准化为十六进制；初始属性采用整块最常见非空属性并补齐
"""

from typing import List, Dict, Any, Optional
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
        if "宋体" in raw or "simsun" in raw:
            return "宋体"
        if "meiryo" in raw:
            return "Meiryo UI"
        # 常见别名/英文字体族名归一
        if raw in {"微软雅黑", "microsoftyahei", "msyahei", "yahei"}:
            return "微软雅黑"
        if raw in {"楷体", "kaiti", "kaitisc", "stkaiti"}:
            return "楷体"
        if raw in {"timesnewroman", "timenewroman", "timesnewromanpsmt"} or value.strip() in {"Times New Roman", "TimeNew Roman", "timenew roman"}:
            return "timenew roman"
        # 若为主题占位符（+mn-ea/+mj-lt 等），归为"其他"（保持不猜测）
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


def _build_initial_attrs(runs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """以首个非空 run 作为初始属性基线（按 run 维度）。"""
    if not runs:
        return {k: None for k in ATTR_KEYS}
    first = next((r for r in runs if str(r.get("字符内容", "")).strip() != ""), None)
    if first is None:
        return {k: None for k in ATTR_KEYS}
    return {k: first.get(k) for k in ATTR_KEYS}


def _emit_text_with_newline_markers(text: str) -> str:
    """将 run 文本中的换行与缩进标记化，同时原样保留其它文本。"""
    if not text:
        return ""
    parts: List[str] = []
    lines = text.split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            parts.append("【换行】")
            # 统计换行后开头空格数以作为缩进
            indent = 0
            for ch in line:
                if ch == ' ':
                    indent += 1
                else:
                    break
            if indent > 0:
                parts.append(f"【缩进{{{indent}}}】")
        parts.append(line)
    return "".join(parts)


def serialize_text_block_to_diff_string(text_block: Dict[str, Any], initial_label: str = "初始的字符所有属性") -> str:
    """按 run 维度进行序列化：
    - 初次输出完整属性【初始的字符所有属性】
    - 当某个 run 的属性相对初始有变化时，输出【字符属性变更：...】（仅列变更项）
    - 文本输出以 run 为单位，run 内部的换行用【换行】与【缩进{N}】标记
    """
    # 取出 runs（原“字符属性”数组，将其视为 run 列表）
    if "字符属性" in text_block:
        runs = text_block["字符属性"]
    else:
        text_info = next(iter(text_block.values()))
        runs = text_info.get("字符属性", [])

    # 预合并：若相邻 run 的属性（除字符内容外）完全一致，则合并为一个 run
    def _norm_run_attrs(r: Dict[str, Any]) -> Dict[str, Any]:
        return {k: _normalize_value(k, r.get(k)) for k in ATTR_KEYS}

    merged_runs: List[Dict[str, Any]] = []
    for r in runs:
        if not merged_runs:
            merged_runs.append(dict(r))
            continue
        last = merged_runs[-1]
        if _norm_run_attrs(last) == _norm_run_attrs(r):
            # 合并文本内容
            last_text = str(last.get("字符内容", ""))
            curr_text = str(r.get("字符内容", ""))
            last["字符内容"] = f"{last_text}{curr_text}"
        else:
            merged_runs.append(dict(r))

    runs = merged_runs

    if not runs:
        return _make_initial_marker({k: None for k in ATTR_KEYS}, initial_label)

    output_parts: List[str] = []
    # 初始属性来自首个非空 run
    initial_attrs_raw = _build_initial_attrs(runs)
    initial_attrs = {k: _normalize_value(k, initial_attrs_raw.get(k)) for k in ATTR_KEYS}
    output_parts.append(_make_initial_marker(initial_attrs, initial_label))

    # 遍历每个 run，若属性与初始不同则输出变更标记，然后输出该 run 文本（含换行标记）
    for run in runs:
        run_text = str(run.get("字符内容", ""))
        curr_attrs_raw = {k: run.get(k) for k in ATTR_KEYS}
        curr_attrs = {k: _normalize_value(k, curr_attrs_raw.get(k)) for k in ATTR_KEYS}

        marker = _make_changed_attrs_marker(initial_attrs, curr_attrs)
        if marker:
            output_parts.append(marker)

        if run_text:
            output_parts.append(_emit_text_with_newline_markers(run_text))

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
