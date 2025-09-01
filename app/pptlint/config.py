"""
配置加载（对应任务：项目骨架与配置解析；为规则与解析器提供阈值与开关）
"""
from dataclasses import dataclass
from typing import List, Optional, Dict, Any
import yaml
import os


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
    llm_provider: str = "deepseek"      # LLM提供商
    llm_model: str = "deepseek-chat"    # 模型名称
    llm_api_key: str = ""               # API密钥
    llm_endpoint: str = ""              # 自定义端点
    llm_temperature: float = 0.2
    llm_max_tokens: int = 1024

    # 审查维度开关
    review_format: bool = True      # 格式规范审查
    review_logic: bool = True       # 内容逻辑审查
    review_acronyms: bool = True    # 缩略语审查
    review_fluency: bool = True     # 表达流畅性审查

    # 审查规则配置
    rules: Dict[str, bool] = None

    # 报告配置
    report: Dict[str, bool] = None

    # 支持的模型列表
    llm_models: Dict[str, List[str]] = None

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
        
        if self.llm_models is None:
            self.llm_models = {
                "deepseek": ["deepseek-chat", "deepseek-coder"],
                "openai": ["gpt-4", "gpt-3.5-turbo", "gpt-4-turbo"],
                "anthropic": ["claude-3-opus", "claude-3-sonnet", "claude-3-haiku"],
                "local": ["qwen2.5-7b", "llama3.1-8b"]
            }
        
        # 如果API key为空，尝试从环境变量获取
        if not self.llm_api_key:
            env_key = f"{self.llm_provider.upper()}_API_KEY"
            self.llm_api_key = os.getenv(env_key, "")


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
        elif key in ["rules", "report", "llm_models"] and isinstance(value, dict):
            # 对于其他嵌套配置，直接传递
            config_data[key] = value
        elif hasattr(ToolConfig, key):
            # 对于直接属性，直接传递
            config_data[key] = value
    
    cfg = ToolConfig(**config_data)
    return cfg

