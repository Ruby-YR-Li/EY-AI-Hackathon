"""LLM 辅助判断模块 — 支持 DeepSeek / 通义千问等 OpenAI 兼容 API。"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any

# ── 模型预设表 ──────────────────────────────────────────────────
# 所有国内大模型统一走 OpenAI 兼容 API，只需换 base_url + model

ModelPreset = dict[str, dict[str, str]]

MODEL_PRESETS: ModelPreset = {
    # ── DeepSeek 系列 ──
    "deepseek-v4-flash": {
        "label": "DeepSeek V4 Flash（推荐·快速）",
        "model": "deepseek-v4-flash",
        "base_url": "https://api.deepseek.com",
        "provider": "deepseek",
    },
    "deepseek-v4-pro": {
        "label": "DeepSeek V4 Pro（高质量）",
        "model": "deepseek-v4-pro",
        "base_url": "https://api.deepseek.com",
        "provider": "deepseek",
    },
    "deepseek-chat": {
        "label": "DeepSeek V3",
        "model": "deepseek-chat",
        "base_url": "https://api.deepseek.com",
        "provider": "deepseek",
    },
    "deepseek-reasoner": {
        "label": "DeepSeek R1（深度推理）",
        "model": "deepseek-reasoner",
        "base_url": "https://api.deepseek.com",
        "provider": "deepseek",
    },
    # ── 通义千问 系列 ──
    "qwen-turbo": {
        "label": "通义千问 Turbo（快速）",
        "model": "qwen-turbo",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "provider": "qwen",
    },
    "qwen-plus": {
        "label": "通义千问 Plus（均衡）",
        "model": "qwen-plus",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "provider": "qwen",
    },
    "qwen-max": {
        "label": "通义千问 Max（最强）",
        "model": "qwen-max",
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "provider": "qwen",
    },
}

DEFAULT_MODEL_KEY = "deepseek-v4-flash"

# 各提供商的默认环境变量名
PROVIDER_ENV_KEYS: dict[str, list[str]] = {
    "deepseek": ["DEEPSEEK_API_KEY"],
    "qwen": ["DASHSCOPE_API_KEY"],
}


def list_models() -> list[dict[str, str]]:
    """返回可用模型列表（供 UI 下拉框使用）。"""
    return [
        {"key": key, "label": info["label"]}
        for key, info in MODEL_PRESETS.items()
    ]


def get_preset(model_key: str | None) -> dict[str, str]:
    """根据 key 获取模型配置，无效 key 返回默认。"""
    if model_key and model_key in MODEL_PRESETS:
        return MODEL_PRESETS[model_key]
    return MODEL_PRESETS[DEFAULT_MODEL_KEY]


@dataclass
class LLMConfig:
    enabled: bool = True
    api_key: str = ""
    model_key: str = DEFAULT_MODEL_KEY
    timeout_seconds: int = 20
    # 以下为派生字段，不需手动设置
    _model: str = field(default="", repr=False)
    _base_url: str = field(default="", repr=False)
    _provider: str = field(default="", repr=False)

    def __post_init__(self):
        preset = get_preset(self.model_key)
        self._model = preset["model"]
        self._base_url = preset["base_url"]
        self._provider = preset.get("provider", "")

    @property
    def model(self) -> str:
        return self._model

    @property
    def base_url(self) -> str:
        return self._base_url

    @property
    def provider(self) -> str:
        return self._provider

    @property
    def effective_api_key(self) -> str:
        if self.api_key.strip():
            return self.api_key.strip()
        for env_key in PROVIDER_ENV_KEYS.get(self._provider, []):
            val = os.environ.get(env_key, "").strip()
            if val:
                return val
        # 兜底：尝试所有已知环境变量
        for keys in PROVIDER_ENV_KEYS.values():
            for k in keys:
                val = os.environ.get(k, "").strip()
                if val:
                    return val
        return ""

    @property
    def is_available(self) -> bool:
        return self.enabled and bool(self.effective_api_key)


def _endpoint(base_url: str) -> str:
    base = (base_url or MODEL_PRESETS[DEFAULT_MODEL_KEY]["base_url"]).rstrip("/")
    if base.endswith("/v1"):
        return f"{base}/chat/completions"
    return f"{base}/v1/chat/completions"


def call_deepseek_json(
    *,
    config: LLMConfig,
    task: str,
    candidates: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], str]:
    if not config.is_available:
        return [], "LLM not configured"
    if not candidates:
        return [], "No candidates"

    system = (
        "你是审计底稿质检助手。只能基于用户提供的结构化证据判断，"
        "不得编造文件名、Sheet、行号、金额或事实。"
        "请仅返回合法 JSON，不要输出 Markdown。"
    )
    user = {
        "task": task,
        "rules": [
            "没有足够证据时 should_raise_note=false",
            "每个结论必须绑定 candidate_id",
            "confidence 范围为 0 到 1",
            "risk_level 只能是 High、Medium、Low",
            "不得新增证据中不存在的位置或金额",
        ],
        "return_schema": {
            "items": [
                {
                    "candidate_id": "string",
                    "should_raise_note": True,
                    "risk_level": "Medium",
                    "issue_type": "string",
                    "reason": "string",
                    "suggested_action": "string",
                    "confidence": 0.75,
                }
            ]
        },
        "candidates": candidates,
    }

    payload = {
        "model": config.model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        _endpoint(config.base_url),
        data=data,
        headers={
            "Authorization": f"Bearer {config.effective_api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=config.timeout_seconds) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:300]
        return [], f"LLM HTTP {exc.code}: {body}"
    except Exception as exc:
        return [], f"LLM call failed: {exc}"

    try:
        content = json.loads(raw)["choices"][0]["message"]["content"]
        parsed = json.loads(content)
        items = parsed.get("items", [])
        if not isinstance(items, list):
            return [], "LLM JSON missing items list"
        return [item for item in items if isinstance(item, dict)], "LLM ok"
    except Exception as exc:
        return [], f"LLM JSON parse failed: {exc}"


def confidence_text(value: object) -> str:
    try:
        return f"{float(value):.2f}"
    except Exception:
        return ""


def test_connection(config: LLMConfig) -> dict:
    """测试 DeepSeek API 连通性。

    发送一个极轻量的 ping 请求（max_tokens=1），返回连通状态和延迟。

    Returns:
        {"ok": bool, "latency_ms": float, "model": str, "error": str}
    """
    import time

    result = {"ok": False, "latency_ms": 0.0, "model": config.model, "error": ""}

    if not config.is_available:
        result["error"] = "API Key 未配置"
        return result

    payload = {
        "model": config.model,
        "messages": [{"role": "user", "content": "ping"}],
        "max_tokens": 1,
        "temperature": 0,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        _endpoint(config.base_url),
        data=data,
        headers={
            "Authorization": f"Bearer {config.effective_api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        t0 = time.perf_counter()
        with urllib.request.urlopen(request, timeout=10) as response:
            response.read()
        result["latency_ms"] = round((time.perf_counter() - t0) * 1000, 1)
        result["ok"] = True
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:200]
        result["error"] = f"HTTP {exc.code}: {body}"
    except Exception as exc:
        result["error"] = str(exc)

    return result
