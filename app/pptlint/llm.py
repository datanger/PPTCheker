"""
LLM 模块（默认启用，可降级）。

说明：
- 接口采用简化的适配器风格，优先走 OpenAI-compatible endpoint（如 DeepSeek）。
- 若未配置 API KEY 或请求失败，返回空建议，调用方需自动降级为规则建议。
"""
import os
import json
import ssl
from typing import Any, Dict, List, Optional
import urllib.request


def _resolve_base_url(provider: str, model: Optional[str], explicit_base_url: Optional[str]) -> Optional[str]:
    """根据提供商与模型推断默认 base url（显式值优先）。"""
    if explicit_base_url:
        return explicit_base_url

    provider_lower = (provider or "").lower()
    model_lower = (model or "").lower()

    # 常见提供商默认 base url（Provider 优先）
    if provider_lower == "local":
        # 内网LLM服务默认地址
        return os.getenv("LLM_BASE_URL", "https://192.168.10.173/sdw/chatbot/sysai/v1")
    if provider_lower == "ollama":
        # Ollama 本地服务
        return os.getenv("LLM_BASE_URL", "http://localhost:11434/v1")
    if provider_lower == "deepseek":
        return os.getenv("LLM_BASE_URL", "https://api.deepseek.com/v1")
    if provider_lower == "openai":
        return os.getenv("LLM_BASE_URL", "https://api.openai.com/v1")
    if provider_lower == "anthropic":
        # Anthropic 并非严格 OpenAI 兼容，但此处仍返回其 messages 根路径
        return os.getenv("LLM_BASE_URL", "https://api.anthropic.com/v1")
    if provider_lower in ("kimi", "moonshot"):
        # Kimi (Moonshot) 采用 OpenAI 兼容接口
        return os.getenv("LLM_BASE_URL", "https://api.moonshot.cn/v1")
    if provider_lower in ("bailian", "dashscope", "aliyun"):
        # 阿里云百炼 DashScope 兼容模式
        return os.getenv("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
    
    # 如果没有明确的 provider，则根据模型名称推断
    if "deepseek" in model_lower:
        return os.getenv("LLM_BASE_URL", "https://api.deepseek.com/v1")
    if model_lower.startswith("gpt"):
        return os.getenv("LLM_BASE_URL", "https://api.openai.com/v1")
    if "claude" in model_lower:
        return os.getenv("LLM_BASE_URL", "https://api.anthropic.com/v1")
    if "moonshot" in model_lower:
        return os.getenv("LLM_BASE_URL", "https://api.moonshot.cn/v1")
    if "qwen" in model_lower:
        return os.getenv("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
    return os.getenv("LLM_BASE_URL")


def _resolve_endpoint(provider: str, model: Optional[str], explicit_endpoint: Optional[str], base_url: Optional[str]) -> Optional[str]:
    """推断最终 endpoint；若给出显式 endpoint 则使用之，否则由 base_url + 路径组合。"""
    if explicit_endpoint:
        return explicit_endpoint

    base = base_url or _resolve_base_url(provider, model, None)
    provider_lower = (provider or "").lower()
    model_lower = (model or "").lower()

    if not base:
        return None

    # OpenAI 兼容路径
    if provider_lower in ("deepseek", "openai", "kimi", "moonshot", "bailian", "dashscope", "aliyun", "local") or any(
        k in model_lower for k in ("gpt", "deepseek", "qwen", "moonshot", "llama", "qwen2")
    ):
        return f"{base.rstrip('/')}/chat/completions"

    # Anthropic messages
    if provider_lower == "anthropic" or "claude" in model_lower:
        return f"{base.rstrip('/')}/messages"

    # 兜底：假定 OpenAI 兼容
    return f"{base.rstrip('/')}/chat/completions"


class LLMClient:
    def __init__(self, provider: str = "deepseek", endpoint: Optional[str] = None, 
                 api_key: Optional[str] = None, model: Optional[str] = None,
                 temperature: float = 0.2, max_tokens: int = 1024,
                 use_proxy: bool = False, proxy_url: Optional[str] = None,
                 base_url: Optional[str] = None):
        self.provider = provider
        self.model = model or "deepseek-chat"
        self.base_url = _resolve_base_url(self.provider, self.model, base_url)
        self.endpoint = _resolve_endpoint(self.provider, self.model, endpoint, self.base_url)
        self.temperature = temperature
        self.max_tokens = max_tokens
        
        # 代理配置
        self.use_proxy = use_proxy
        self.proxy_url = proxy_url
        
        # 根据提供商设置API key
        if api_key:
            self.api_key = api_key
        else:
            # 为local和ollama设置默认API key
            provider_lower = provider.lower()
            if provider_lower == "local":
                self.api_key = "local-api-key"
            elif provider_lower == "ollama":
                self.api_key = "ollama-api-key"
            else:
                # 尝试从环境变量获取对应提供商的API key
                env_key = f"{provider.upper()}_API_KEY"
                self.api_key = os.getenv(env_key, "")

    def complete(self, prompt: str, max_tokens: Optional[int] = None, stop_event: Optional[object] = None) -> str:
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
            
            # 检查是否应该停止
            if stop_event and stop_event.is_set():
                print("⏹️ LLM调用被用户终止")
                return ""
            
            try:
                # 配置代理处理
                if self.use_proxy and self.proxy_url:
                    # 启用代理
                    proxy_handler = urllib.request.ProxyHandler({
                        'http': self.proxy_url,
                        'https': self.proxy_url
                    })
                    opener = urllib.request.build_opener(proxy_handler)
                    urllib.request.install_opener(opener)
                    print(f"🌐 使用代理: {self.proxy_url}")
                else:
                    # 禁用代理，清除环境变量影响
                    proxy_handler = urllib.request.ProxyHandler({})
                    opener = urllib.request.build_opener(proxy_handler)
                    urllib.request.install_opener(opener)
                
                # 检查是否需要跳过SSL验证（内网地址）
                context = None
                if self.endpoint and ("192.168." in self.endpoint or "10." in self.endpoint or "172." in self.endpoint):
                    context = ssl.create_default_context()
                    context.check_hostname = False
                    context.verify_mode = ssl.CERT_NONE
                    print(f"🔓 跳过SSL验证: {self.endpoint}")
                
                with urllib.request.urlopen(req, data=data, context=context) as resp:
                    payload = json.loads(resp.read().decode("utf-8"))
                    
                    # 检查是否有错误
                    if "error" in payload and payload["error"]:
                        print(f"LLM API错误: {payload['error'].get('message', '未知错误')}")
                        return ""
                    
                    # OpenAI style - 安全解析
                    choices = payload.get("choices", [])
                    if choices and len(choices) > 0:
                        message = choices[0].get("message", {})
                        return message.get("content", "")
                    
                    print("LLM未返回有效内容")
                    return ""
            except Exception as e:
                print(f"LLM调用异常: {e}")
                return ""
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


