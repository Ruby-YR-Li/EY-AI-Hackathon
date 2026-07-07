"""首次安装程序：安装固定启动器、初始版本，并创建桌面快捷方式。"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import sys
import zipfile
from pathlib import Path
import win_dialogs


APP_HOME = Path(os.getenv("LOCALAPPDATA", Path.home())) / "QualityAgent"


def bundle_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def desktop_dir() -> Path:
    """读取Windows登记的真实桌面路径，兼容企业桌面重定向。"""
    buffer = ctypes.create_unicode_buffer(32768)
    result = ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer)
    if result == 0 and buffer.value:
        return Path(buffer.value)
    return Path(os.getenv("USERPROFILE", Path.home())) / "Desktop"


def create_desktop_launcher(target: Path) -> Path:
    """直接复制固定启动器到桌面，避免依赖PowerShell、WScript和.lnk权限。"""
    desktop = desktop_dir()
    desktop.mkdir(parents=True, exist_ok=True)
    desktop_launcher = desktop / "质检Agent.exe"
    shutil.copy2(target, desktop_launcher)
    return desktop_launcher


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--silent", action="store_true", help="安装验证时不显示完成弹窗")
    args = parser.parse_args()
    try:
        source = bundle_dir()
        manifest = json.loads((source / "install_manifest.json").read_text(encoding="utf-8"))
        version = str(manifest["version"])
        launcher_source = source / "启动质检Agent.exe"
        app_zip = source / "初始版本.zip"
        if not launcher_source.is_file() or not app_zip.is_file():
            raise FileNotFoundError("首次安装包不完整，请重新获取安装包")

        APP_HOME.mkdir(parents=True, exist_ok=True)
        (APP_HOME / "config").mkdir(exist_ok=True)
        versions = APP_HOME / "versions"
        versions.mkdir(exist_ok=True)
        launcher_target = APP_HOME / "启动质检Agent.exe"
        shutil.copy2(launcher_source, launcher_target)

        version_dir = versions / version
        if version_dir.exists():
            shutil.rmtree(version_dir)
        version_dir.mkdir()
        with zipfile.ZipFile(app_zip) as archive:
            archive.extractall(version_dir)
        children = list(version_dir.iterdir())
        if len(children) == 1 and children[0].is_dir():
            nested = children[0]
            for child in nested.iterdir():
                shutil.move(str(child), version_dir / child.name)
            nested.rmdir()
        (APP_HOME / "current.json").write_text(
            json.dumps({"version": version, "path": str(version_dir)}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        desktop_error = ""
        try:
            create_desktop_launcher(launcher_target)
        except OSError as exc:
            # 公司策略禁止写桌面时，主体安装仍然有效，不能误报为安装失败。
            desktop_error = str(exc)
        if not args.silent:
            if desktop_error:
                win_dialogs.warning(
                    "安装完成（未创建桌面图标）",
                    "质检Agent已经安装成功，但公司电脑不允许程序写入桌面。\n\n"
                    f"请双击以下文件启动：\n{launcher_target}\n\n"
                    f"桌面写入错误：{desktop_error}",
                )
            else:
                win_dialogs.info(
                    "安装完成",
                    "质检Agent已安装完成。\n\n桌面已生成“质检Agent”启动程序，"
                    "以后只需双击该图标即可自动更新并启动。",
                )
    except Exception as exc:
        if not args.silent:
            win_dialogs.error("安装失败", str(exc))
        raise


if __name__ == "__main__":
    main()
