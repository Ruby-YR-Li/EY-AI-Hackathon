"""
审计底稿质检Agent - AI调用结果缓存
减少重复API调用，降低成本
"""

import hashlib
import json
import os
import time
from typing import Optional

from .config import AI_CACHE_TTL
from .app_paths import CACHE_DIR as USER_CACHE_DIR

CACHE_DIR = str(USER_CACHE_DIR)


class AICache:
    """AI调用结果文件缓存"""

    def __init__(self, cache_dir: str = CACHE_DIR, ttl: int = AI_CACHE_TTL):
        self.cache_dir = cache_dir
        self.ttl = ttl
        os.makedirs(cache_dir, exist_ok=True)

    def _get_cache_key(self, prompt: str) -> str:
        """生成缓存键（基于提示词内容的MD5）"""
        return hashlib.md5(prompt.encode("utf-8")).hexdigest()

    def _get_cache_path(self, key: str) -> str:
        """获取缓存文件路径"""
        return os.path.join(self.cache_dir, f"{key}.json")

    def get(self, prompt: str) -> Optional[dict]:
        """
        从缓存中获取结果

        Args:
            prompt: 提示词

        Returns:
            dict 或 None（缓存不存在或已过期）
        """
        key = self._get_cache_key(prompt)
        path = self._get_cache_path(key)

        if not os.path.exists(path):
            return None

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # 检查是否过期
            if time.time() - data.get("timestamp", 0) > self.ttl:
                os.remove(path)
                return None

            return data.get("result")
        except (json.JSONDecodeError, IOError):
            return None

    def set(self, prompt: str, result: dict):
        """
        将结果写入缓存

        Args:
            prompt: 提示词
            result: 判断结果
        """
        key = self._get_cache_key(prompt)
        path = self._get_cache_path(key)

        data = {
            "timestamp": time.time(),
            "result": result,
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

    def clear(self):
        """清空所有缓存"""
        for filename in os.listdir(self.cache_dir):
            if filename.endswith(".json"):
                os.remove(os.path.join(self.cache_dir, filename))
