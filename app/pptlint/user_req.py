"""
用户审查需求解析器：从《审查需求文档.md》提取配置（容错，默认回落）。

实现策略：
- 采用简单正则/关键词解析；若提供 YAML 则更精确（预留）。
"""
import re
from typing import Optional

from .config import ToolConfig


def parse_user_requirements(md_path: str, base: ToolConfig) -> ToolConfig:
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception:
        return base

    cfg = ToolConfig(**vars(base))

    m = re.search(r"日文字体名.*?[:：]\s*(.+)", content)
    if m:
        name = m.group(1).strip()
        if name:
            cfg.jp_font_name = name

    m = re.search(r"最小字号.*?(\d+)", content)
    if m:
        try:
            cfg.min_font_size_pt = int(m.group(1))
        except Exception:
            pass

    m = re.search(r"单页颜色上限.*?(\d+)", content)
    if m:
        try:
            cfg.color_count_threshold = int(m.group(1))
        except Exception:
            pass

    # 模式（review/edit）
    if re.search(r"模式.*?编辑|edit", content):
        # 仅记录在cfg中，CLI最终决定
        pass

    return cfg


