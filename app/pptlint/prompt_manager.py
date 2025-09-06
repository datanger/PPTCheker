"""
LLM提示词管理模块

功能：
- 统一管理所有LLM审查的提示词配置
- 支持用户自定义用户提示部分
- 自动注入输入参数和输出格式
- 提供提示词的查看和编辑功能
"""
import os
import yaml
from typing import Dict, Any, List
from dataclasses import dataclass


@dataclass
class PromptTemplate:
    """提示词模板类"""
    name: str
    description: str
    user_prompt: str
    
    def get_user_prompt(self, **kwargs) -> str:
        """获取格式化的用户提示"""
        return self.user_prompt.format(**kwargs)


class PromptManager:
    """提示词管理器"""
    
    def __init__(self, config_path: str = None):
        self.config_path = config_path or self._get_default_config_path()
        self.prompts: Dict[str, PromptTemplate] = {}
        self.load_prompts()
    
    def _get_default_config_path(self) -> str:
        """获取默认配置文件路径"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, "..", "configs", "llm_prompts.yaml")
    
    def load_prompts(self):
        """加载提示词配置"""
        try:
            if not os.path.exists(self.config_path):
                print(f"⚠️ 提示词配置文件不存在: {self.config_path}")
                return
            
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            
            llm_prompts = config.get('llm_prompts', {})
            
            for key, prompt_data in llm_prompts.items():
                self.prompts[key] = PromptTemplate(
                    name=prompt_data.get('name', ''),
                    description=prompt_data.get('description', ''),
                    user_prompt=prompt_data.get('user_prompt', '')
                )
            
            print(f"✅ 加载了 {len(self.prompts)} 个提示词配置")
            
        except Exception as e:
            print(f"❌ 加载提示词配置失败: {e}")
    
    def save_prompts(self):
        """保存提示词配置"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            
            config = {
                'llm_prompts': {}
            }
            
            for key, prompt in self.prompts.items():
                config['llm_prompts'][key] = {
                    'name': prompt.name,
                    'description': prompt.description,
                    'user_prompt': prompt.user_prompt
                }
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, ensure_ascii=False, indent=2, allow_unicode=True)
            
            print(f"✅ 提示词配置已保存到: {self.config_path}")
            return True
            
        except Exception as e:
            print(f"❌ 保存提示词配置失败: {e}")
            return False
    
    def get_prompt(self, key: str) -> PromptTemplate:
        """获取指定提示词模板"""
        return self.prompts.get(key)
    
    def get_all_prompts(self) -> Dict[str, PromptTemplate]:
        """获取所有提示词模板"""
        return self.prompts.copy()
    
    def update_user_prompt(self, key: str, user_prompt: str):
        """更新用户提示部分"""
        if key in self.prompts:
            self.prompts[key].user_prompt = user_prompt
            print(f"✅ 已更新 {key} 的用户提示")
        else:
            print(f"❌ 未找到提示词: {key}")
    
    def get_prompt_names(self) -> List[str]:
        """获取所有提示词名称列表"""
        return list(self.prompts.keys())
    
    def get_prompt_info(self, key: str) -> Dict[str, str]:
        """获取提示词信息（不包含完整内容）"""
        if key in self.prompts:
            prompt = self.prompts[key]
            return {
                'name': prompt.name,
                'description': prompt.description,
                'user_prompt_preview': prompt.user_prompt[:100] + "..." if len(prompt.user_prompt) > 100 else prompt.user_prompt
            }
        return {}
    
    def get_user_prompt_for_review(self, review_type: str, **kwargs) -> str:
        """获取特定审查类型的用户提示词"""
        prompt_template = self.get_prompt(review_type)
        if prompt_template:
            return prompt_template.get_user_prompt(**kwargs)
        else:
            print(f"❌ 未找到审查类型: {review_type}")
            return ""


# 全局提示词管理器实例
prompt_manager = PromptManager()
