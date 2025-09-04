"""
LLM 模块（默认启用，可降级）。

说明：
- 接口采用简化的适配器风格，优先走 OpenAI-compatible endpoint（如 DeepSeek）。
- 若未配置 API KEY 或请求失败，返回空建议，调用方需自动降级为规则建议。
"""
import os
import json
from typing import Any, Dict, List, Optional
import urllib.request


def _resolve_endpoint(provider: str, model: Optional[str], explicit_endpoint: Optional[str]) -> Optional[str]:
    """根据提供商和模型名推断默认 endpoint（OpenAI-compatible），显式值优先。"""
    if explicit_endpoint:
        return explicit_endpoint
    
    provider_lower = provider.lower()
    model_lower = (model or "").lower()
    
    # 常见提供方的默认 OpenAI-compatible chat.completions 端点
    if provider_lower == "deepseek" or "deepseek" in model_lower:
        return os.getenv("LLM_ENDPOINT", "https://api.deepseek.com/v1/chat/completions")
    elif provider_lower == "openai" or model_lower.startswith("gpt"):
        return os.getenv("LLM_ENDPOINT", "https://api.openai.com/v1/chat/completions")
    elif provider_lower == "anthropic" or "claude" in model_lower:
        return os.getenv("LLM_ENDPOINT", "https://api.anthropic.com/v1/messages")
    elif provider_lower == "local":
        return os.getenv("LLM_ENDPOINT", "http://localhost:11434/v1/chat/completions")  # Ollama默认端点
    else:
        # 兜底用环境变量
        return os.getenv("LLM_ENDPOINT")


class LLMClient:
    def __init__(self, provider: str = "deepseek", endpoint: Optional[str] = None, 
                 api_key: Optional[str] = None, model: Optional[str] = None,
                 temperature: float = 0.2, max_tokens: int = 1024):
        self.provider = provider
        self.model = model or "deepseek-chat"
        self.endpoint = _resolve_endpoint(self.provider, self.model, endpoint)
        self.temperature = temperature
        self.max_tokens = max_tokens
        
        # 根据提供商设置API key
        if api_key:
            self.api_key = api_key
        else:
            # 尝试从环境变量获取对应提供商的API key
            env_key = f"{provider.upper()}_API_KEY"
            self.api_key = os.getenv(env_key, "")

    def complete(self, prompt: str, max_tokens: Optional[int] = None) -> str:
        try:
            if not self.api_key:
                print("未配置API密钥，LLM功能将不可用")
                return ""
            
            req = urllib.request.Request(self.endpoint, method="POST")
            req.add_header("Content-Type", "application/json")
            req.add_header("Authorization", f"Bearer {self.api_key}")
            
            # 每次调用都使用新的对话上下文，避免历史对话干扰
            body = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": "You are a helpful assistant for document review."},
                    {"role": "user", "content": prompt},
                ],
                "max_tokens": max_tokens or self.max_tokens,
                "temperature": self.temperature,
            }
            
            data = json.dumps(body).encode("utf-8")
            
            # 增加超时时间，并添加重试机制
            timeout = 60  # 增加到60秒
            max_retries = 2
            
            for attempt in range(max_retries + 1):
                try:
                    with urllib.request.urlopen(req, data=data, timeout=timeout) as resp:
                        payload = json.loads(resp.read().decode("utf-8"))
                        # OpenAI style
                        return payload.get("choices", [{}])[0].get("message", {}).get("content", "")
                except urllib.error.URLError as e:
                    if "timeout" in str(e).lower() and attempt < max_retries:
                        print(f"LLM请求超时，第{attempt + 1}次重试...")
                        continue
                    else:
                        raise e
        except Exception as e:
            print(f"LLM调用异常: {e}")
            return ""


def suggest_japanese_fluency(llm: LLMClient, text: str, constraints: str = "") -> List[str]:
    prompt = f"改写为自然流畅的日本汽车IT行业表述，保持技术准确性：\n约束:{constraints}\n文本:\n{text}"
    out = llm.complete(prompt)
    return [s.strip() for s in out.splitlines() if s.strip()] if out else []


def suggest_logic_transition(llm: LLMClient, outline: str) -> List[str]:
    prompt = f"为以下PPT大纲提出过渡与连贯性建议（简短要点）：\n{outline}"
    out = llm.complete(prompt)
    return [s.strip() for s in out.splitlines() if s.strip()] if out else []


def suggest_term_unification(llm: LLMClient, variants: List[str]) -> Optional[str]:
    prompt = "请在以下术语变体中选择统一用法（只输出一个最佳写法）：\n" + "\n".join(variants)
    out = llm.complete(prompt)
    return out.strip() if out else None


