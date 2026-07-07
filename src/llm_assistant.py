"""DeepSeek-assisted evidence judgement for the expense review MVP."""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any


DEFAULT_DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEFAULT_DEEPSEEK_MODEL = "deepseek-chat"


@dataclass
class LLMConfig:
    enabled: bool = True
    api_key: str = ""
    base_url: str = DEFAULT_DEEPSEEK_BASE_URL
    model: str = DEFAULT_DEEPSEEK_MODEL
    timeout_seconds: int = 20

    @property
    def effective_api_key(self) -> str:
        return self.api_key.strip() or os.environ.get("DEEPSEEK_API_KEY", "").strip() or os.environ.get("OPENAI_API_KEY", "").strip()

    @property
    def is_available(self) -> bool:
        return self.enabled and bool(self.effective_api_key)


def _endpoint(base_url: str) -> str:
    base = (base_url or DEFAULT_DEEPSEEK_BASE_URL).rstrip("/")
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
        "model": config.model or DEFAULT_DEEPSEEK_MODEL,
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

    result = {"ok": False, "latency_ms": 0.0, "model": config.model or DEFAULT_DEEPSEEK_MODEL, "error": ""}

    if not config.is_available:
        result["error"] = "API Key 未配置"
        return result

    payload = {
        "model": config.model or DEFAULT_DEEPSEEK_MODEL,
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
