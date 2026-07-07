"""
审计底稿质检Agent - API配置
支持从环境变量读取配置
"""

import os


def _get_env(name: str, default: str = "") -> str:
    return os.getenv(name, default)


AI_PROVIDER = _get_env("AI_PROVIDER", "auto")           # auto / anthropic / openai / qwen
AI_API_KEY = _get_env("AI_API_KEY", "")
AI_MODEL = _get_env("AI_MODEL", "claude-sonnet-4-20250514")
AI_BASE_URL = _get_env("AI_BASE_URL", "")             # 自定义API端点（可选）
ENABLE_AI_CHECK = _get_env("ENABLE_AI_CHECK", "true").lower() == "true"
AI_CACHE_TTL = int(_get_env("AI_CACHE_TTL", "86400"))  # 缓存时间（秒），默认24小时
AI_MAX_RETRIES = int(_get_env("AI_MAX_RETRIES", "2"))  # API调用最大重试次数
AI_REQUEST_TIMEOUT = float(_get_env("AI_REQUEST_TIMEOUT", "60"))  # 单次AI请求超时（秒）


def reload_config():
    """重新从环境变量加载配置"""
    global AI_PROVIDER, AI_API_KEY, AI_MODEL, AI_BASE_URL, ENABLE_AI_CHECK, AI_CACHE_TTL, AI_MAX_RETRIES, AI_REQUEST_TIMEOUT
    AI_PROVIDER = _get_env("AI_PROVIDER", "auto")
    AI_API_KEY = _get_env("AI_API_KEY", "")
    AI_MODEL = _get_env("AI_MODEL", "claude-sonnet-4-20250514")
    AI_BASE_URL = _get_env("AI_BASE_URL", "")
    ENABLE_AI_CHECK = _get_env("ENABLE_AI_CHECK", "true").lower() == "true"
    AI_CACHE_TTL = int(_get_env("AI_CACHE_TTL", "86400"))
    AI_MAX_RETRIES = int(_get_env("AI_MAX_RETRIES", "2"))
    AI_REQUEST_TIMEOUT = float(_get_env("AI_REQUEST_TIMEOUT", "60"))
