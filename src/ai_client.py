"""
审计底稿质检Agent - 大模型API客户端封装
支持多种大模型API：Anthropic Claude、OpenAI、通义千问
"""

import json
import time
from abc import ABC, abstractmethod
from typing import Optional

from .cache import AICache
from .config import (
    AI_PROVIDER,
    AI_API_KEY,
    AI_MODEL,
    AI_BASE_URL,
    AI_MAX_RETRIES,
    AI_REQUEST_TIMEOUT,
)


class AIJudgeResult:
    """大模型判断结果"""

    def __init__(self, has_issue: bool, reason: str, suggestion: str = "",
                 is_error: bool = False):
        self.has_issue = has_issue      # 是否存在问题
        self.reason = reason            # 判断理由
        self.suggestion = suggestion    # 改进建议
        self.is_error = is_error        # 是否为调用失败的错误结果

    def to_dict(self) -> dict:
        return {"has_issue": self.has_issue, "reason": self.reason,
                "suggestion": self.suggestion, "is_error": self.is_error}

    @classmethod
    def from_dict(cls, d: dict) -> "AIJudgeResult":
        return cls(
            has_issue=d.get("has_issue", False),
            reason=d.get("reason", ""),
            suggestion=d.get("suggestion", ""),
            is_error=d.get("is_error", False)
        )


class BaseAIClient(ABC):
    """大模型API客户端基类"""

    def __init__(self, api_key: str, model: str, base_url: Optional[str] = None):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.cache = AICache()

    @abstractmethod
    def call_api(self, prompt: str) -> str:
        """调用大模型API，返回原始响应文本"""
        pass

    def judge(self, prompt: str, use_cache: bool = True) -> AIJudgeResult:
        """
        调用大模型进行判断

        Args:
            prompt: 提示词
            use_cache: 是否使用缓存

        Returns:
            AIJudgeResult: 判断结果
        """
        # 检查缓存
        if use_cache:
            cached = self.cache.get(prompt)
            if cached:
                cached_result = AIJudgeResult.from_dict(cached)
                # 历史版本曾缓存调用/解析失败结果，导致后续检查持续复用错误。
                # 错误缓存一律忽略，重新调用模型。
                if not cached_result.is_error:
                    return cached_result

        # 调用API并解析（带重试）。空响应和JSON解析失败也属于可重试错误。
        last_error = None
        last_response = ""
        for attempt in range(AI_MAX_RETRIES + 1):
            try:
                response_text = self.call_api(prompt)
                last_response = "" if response_text is None else str(response_text)
                if not last_response.strip():
                    raise ValueError("模型返回空响应")
                result = self._parse_response(last_response)
                if use_cache:
                    self.cache.set(prompt, result.to_dict())
                return result
            except Exception as e:
                last_error = e
                if attempt == AI_MAX_RETRIES:
                    return AIJudgeResult(
                        has_issue=True,
                        reason=(
                            f"AI调用或响应解析失败（{attempt+1}次尝试后）: {str(e)}"
                            + (
                                f"，原始响应: {last_response[:200]}"
                                if last_response.strip() else ""
                            )
                        ),
                        suggestion="建议人工复核此检查点",
                        is_error=True
                    )
                time.sleep(2 ** attempt)  # 指数退避

        return AIJudgeResult(
            has_issue=True,
            reason=f"AI调用失败: {last_error}",
            suggestion="建议人工复核此检查点",
            is_error=True,
        )

    def _parse_response(self, response_text: str) -> AIJudgeResult:
        """
        解析大模型响应，提取JSON格式的判断结果

        支持两种格式：
        1. 纯JSON：直接解析
        2. JSON在markdown代码块中：先提取再解析
        """
        text = response_text.strip()

        # 尝试从markdown代码块中提取JSON
        if "```json" in text:
            start = text.find("```json") + 7
            end = text.find("```", start)
            text = text[start:end].strip()
        elif "```" in text:
            start = text.find("```") + 3
            end = text.find("```", start)
            text = text[start:end].strip()

        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            # 兼容模型在JSON前后附带少量说明文字的情况。
            start = text.find("{")
            end = text.rfind("}")
            if start < 0 or end <= start:
                raise
            data = json.loads(text[start:end + 1])
        if not isinstance(data, dict) or "has_issue" not in data:
            raise ValueError("AI响应JSON缺少has_issue字段")
        return AIJudgeResult.from_dict(data)


class AnthropicClient(BaseAIClient):
    """Anthropic Claude API"""

    def call_api(self, prompt: str) -> str:
        try:
            import anthropic
        except ImportError:
            raise ImportError("请安装 anthropic SDK: pip install anthropic")

        client = anthropic.Anthropic(
            api_key=self.api_key,
            base_url=self.base_url or None,
            timeout=AI_REQUEST_TIMEOUT,
        )

        message = client.messages.create(
            model=self.model,
            max_tokens=4096,
            temperature=0,
            messages=[{"role": "user", "content": prompt}]
        )

        return message.content[0].text


class OpenAIClient(BaseAIClient):
    """OpenAI API"""

    def call_api(self, prompt: str) -> str:
        try:
            import openai
        except ImportError:
            raise ImportError("请安装 openai SDK: pip install openai")

        client = openai.OpenAI(
            api_key=self.api_key,
            base_url=self.base_url or None,
            timeout=AI_REQUEST_TIMEOUT,
        )

        response = client.chat.completions.create(
            model=self.model,
            max_tokens=4096,
            temperature=0,
            top_p=1,
            messages=[{"role": "user", "content": prompt}]
        )

        return response.choices[0].message.content


class QwenClient(BaseAIClient):
    """通义千问 API（兼容OpenAI接口格式）"""

    def call_api(self, prompt: str) -> str:
        try:
            import openai
        except ImportError:
            raise ImportError("请安装 openai SDK: pip install openai")

        base_url = self.base_url or "https://dashscope.aliyuncs.com/compatible-mode/v1"
        client = openai.OpenAI(
            api_key=self.api_key,
            base_url=base_url,
            timeout=AI_REQUEST_TIMEOUT,
        )

        response = client.chat.completions.create(
            model=self.model,
            max_tokens=4096,
            temperature=0,
            top_p=1,
            messages=[{"role": "user", "content": prompt}]
        )

        return response.choices[0].message.content


def _detect_provider_from_key(api_key: str) -> str:
    """根据API密钥特征自动识别供应商"""
    if api_key.startswith('sk-ant-'):
        return 'anthropic'
    # 所有以sk-开头的都用OpenAI兼容接口（覆盖deepseek、qwen等国内厂商）
    if api_key.startswith('sk-'):
        return 'openai'
    return 'openai'


def create_ai_client() -> Optional[BaseAIClient]:
    """
    根据配置创建AI客户端

    Returns:
        BaseAIClient 或 None（如果未配置API密钥）
    """
    if not AI_API_KEY:
        return None

    clients = {
        "anthropic": AnthropicClient,
        "openai": OpenAIClient,
        "qwen": QwenClient,
    }

    # 自动检测供应商：如果provider为auto或未匹配，根据密钥特征自动识别
    provider = AI_PROVIDER
    if provider == 'auto' or provider not in clients:
        provider = _detect_provider_from_key(AI_API_KEY)

    client_class = clients.get(provider)
    if client_class is None:
        raise ValueError(f"无法识别AI服务供应商，请检查API密钥格式。支持的供应商: {list(clients.keys())}")

    return client_class(
        api_key=AI_API_KEY,
        model=AI_MODEL,
        base_url=AI_BASE_URL
    )
