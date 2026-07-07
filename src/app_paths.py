"""应用资源目录与每位用户的本地数据目录。"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path


APP_NAME = "QualityAgent"
DISPLAY_NAME = "质检Agent"
try:
    from ._build_version import APP_VERSION
except ImportError:
    APP_VERSION = "dev"


def resource_dir() -> Path:
    """源码运行时返回项目根目录，打包后返回PyInstaller资源目录。"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)).resolve()
    return Path(__file__).resolve().parent.parent


def user_data_dir() -> Path:
    override = os.getenv("QUALITY_AGENT_DATA_DIR", "").strip()
    if override:
        return Path(override).expanduser().resolve()
    local_app_data = os.getenv("LOCALAPPDATA")
    candidate = Path(local_app_data) / APP_NAME if local_app_data else Path.home() / f".{APP_NAME}"
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        return candidate
    except OSError:
        # 受限开发/测试环境可能禁止写LOCALAPPDATA；正式安装环境仍优先使用上述目录。
        fallback = resource_dir() / ".local_data"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


ROOT_DIR = user_data_dir()
CONFIG_DIR = ROOT_DIR / "config"
UPLOAD_DIR = ROOT_DIR / "待质检底稿"
OUTPUT_DIR = ROOT_DIR / "质检结果"
CACHE_DIR = ROOT_DIR / "cache"
LOG_DIR = ROOT_DIR / "logs"
VERSIONS_DIR = ROOT_DIR / "versions"
DOWNLOAD_DIR = ROOT_DIR / "updates"

AI_CONFIG_FILE = CONFIG_DIR / "ai_config.json"
LAUNCHER_CONFIG_FILE = CONFIG_DIR / "launcher_config.json"
PROJECT_CONFIG_FILE = CONFIG_DIR / "projects.json"
CURRENT_VERSION_FILE = ROOT_DIR / "current.json"


def ensure_user_dirs() -> None:
    for path in (
        ROOT_DIR,
        CONFIG_DIR,
        UPLOAD_DIR,
        OUTPUT_DIR,
        CACHE_DIR,
        LOG_DIR,
        VERSIONS_DIR,
        DOWNLOAD_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)


def _copy_if_missing(source: Path, target: Path) -> bool:
    if target.exists() or not source.exists() or not source.is_file():
        return False
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    return True


def migrate_legacy_data() -> list[str]:
    """首次运行时迁移旧版放在程序目录内的配置，不移动真实底稿和报告。"""
    ensure_user_dirs()
    base = resource_dir()
    migrated = []
    candidates = [
        (base / "ai_config.json", AI_CONFIG_FILE),
        (base / "资料库" / "项目配置" / "projects.json", PROJECT_CONFIG_FILE),
    ]
    for source, target in candidates:
        if _copy_if_missing(source, target):
            migrated.append(str(target))
    if not PROJECT_CONFIG_FILE.exists():
        PROJECT_CONFIG_FILE.write_text("{}", encoding="utf-8")
    return migrated


def load_json(path: Path, default=None):
    path = Path(path)
    if default is None:
        default = {}
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        pass
    return default.copy() if isinstance(default, dict) else default


def save_json(path: Path, data) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


ensure_user_dirs()
migrate_legacy_data()
