"""
配置加载（对应任务：项目骨架与配置解析；为规则与解析器提供阈值与开关）
"""
from dataclasses import dataclass
from typing import List, Optional, Dict, Any
import yaml


@dataclass
class ToolConfig:
    # 字体/字号
    jp_font_name: str = "Meiryo UI"  # 日文字体统一
    min_font_size_pt: int = 12

    # 缩略语（由LLM智能识别，无需手动配置）
    # 注意：缩略语识别完全由LLM大模型进行

    # 颜色
    color_count_threshold: int = 5

    # 输出格式
    output_format: str = "md"  # md | html

    # 自动修复白名单
    autofix_font: bool = False
    autofix_size: bool = False
    autofix_color: bool = False

    # 词库路径（可选）
    jp_terms_path: Optional[str] = None
    term_mapping_path: Optional[str] = None

    # LLM配置
    llm_enabled: bool = True
    llm_provider: str = "deepseek"      # LLM提供商：deepseek, openai, anthropic, local
    llm_model: str = "deepseek-chat"
    llm_api_key: Optional[str] = None   # API密钥
    llm_endpoint: Optional[str] = None  # 自定义端点（留空则使用默认）
    llm_temperature: float = 0.2
    llm_max_tokens: int = 9999
    llm_use_proxy: bool = False         # 是否使用代理（默认关闭）
    llm_proxy_url: Optional[str] = None # 代理URL

    # 审查维度开关
    review_format: bool = True      # 格式规范审查
    review_logic: bool = True       # 内容逻辑审查
    review_acronyms: bool = True    # 缩略语审查
    review_fluency: bool = True     # 表达流畅性审查

    # 审查规则配置
    rules: Dict[str, bool] = None

    # 报告配置
    report: Dict[str, bool] = None

    def __post_init__(self):
        # 设置默认值
        if self.rules is None:
            self.rules = {
                "font_family": True,
                "font_size": True,
                "color_count": True,
                "theme_harmony": True,
                "acronym_explanation": True
            }
        
        if self.report is None:
            self.report = {
                "include_summary": True,
                "include_details": True,
                "include_suggestions": True,
                "include_statistics": True
            }


def load_config(path: str) -> ToolConfig:
    """从 YAML 加载配置。"""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    
    # 处理嵌套配置
    config_data = {}
    for key, value in data.items():
        if key == "llm_review" and isinstance(value, dict):
            # 处理llm_review嵌套配置
            for review_key, review_value in value.items():
                if hasattr(ToolConfig, review_key):
                    config_data[review_key] = review_value
        elif key == "rules_review" and isinstance(value, dict):
            # 处理rules_review嵌套配置，映射到rules
            config_data["rules"] = value
        elif key in ["rules", "report"] and isinstance(value, dict):
            # 对于其他嵌套配置，直接传递
            config_data[key] = value
        elif hasattr(ToolConfig, key):
            # 对于直接属性，直接传递
            config_data[key] = value
    
    cfg = ToolConfig(**config_data)
    return cfg

