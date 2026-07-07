from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any


class LLMUnavailable(RuntimeError):
    pass


@dataclass
class LLMConfig:
    api_key: str
    base_url: str
    model: str
    timeout: int
    retries: int


def get_llm_config() -> LLMConfig:
    api_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise LLMUnavailable("未设置 DEEPSEEK_API_KEY，已回退到规则结果。")
    return LLMConfig(
        api_key=api_key,
        base_url=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com/chat/completions").strip(),
        model=os.environ.get("DEEPSEEK_MODEL", "deepseek-chat").strip(),
        timeout=int(os.environ.get("DEEPSEEK_TIMEOUT", "45")),
        retries=int(os.environ.get("DEEPSEEK_RETRIES", "1")),
    )


def chat_json(messages: list[dict[str, str]], *, temperature: float = 0.0, max_tokens: int = 4096) -> Any:
    cfg = get_llm_config()
    payload = {
        "model": cfg.model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    last_error: Exception | None = None
    for attempt in range(cfg.retries + 1):
        req = urllib.request.Request(
            cfg.base_url,
            data=data,
            headers={
                "Authorization": f"Bearer {cfg.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=cfg.timeout) as resp:
                body = json.loads(resp.read().decode("utf-8"))
            content = body["choices"][0]["message"]["content"]
            return extract_json(content)
        except (urllib.error.URLError, urllib.error.HTTPError, KeyError, json.JSONDecodeError, ValueError) as exc:
            last_error = exc
            if attempt < cfg.retries:
                time.sleep(0.8 * (attempt + 1))
    raise LLMUnavailable(f"LLM 调用失败，已回退到规则结果：{last_error}")


def extract_json(text: str) -> Any:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"(\{.*\}|\[.*\])", text, flags=re.DOTALL)
        if not match:
            raise
        return json.loads(match.group(1))
