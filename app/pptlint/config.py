"""
配置加载（对应任务：项目骨架与配置解析；为规则与解析器提供阈值与开关）
"""
from dataclasses import dataclass
from typing import List, Optional
import yaml


@dataclass
class ToolConfig:
    # 字体/字号
    jp_font_name: str = "Meiyou UI"  # 日文字体统一
    min_font_size_pt: int = 12

    # 缩略语
    acronym_min_len: int = 2
    acronym_max_len: int = 8

    # 颜色
    color_count_threshold: int = 5

    # 报告
    output_format: str = "md"  # md | html

    # 自动修复白名单
    autofix_font: bool = False
    autofix_size: bool = False
    autofix_color: bool = False

    # 词库路径（可选）
    jp_terms_path: Optional[str] = None
    term_mapping_path: Optional[str] = None


def load_config(path: str) -> ToolConfig:
    """从 YAML 加载配置。"""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cfg = ToolConfig(**{k: v for k, v in data.items() if hasattr(ToolConfig, k)})
    return cfg

