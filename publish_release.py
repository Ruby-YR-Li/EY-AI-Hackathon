"""内网一键发布工具：把新版程序包放入共享文件夹并更新latest.json。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from pathlib import Path
import win_dialogs


PUBLISHER_CONFIG = (
    Path(os.getenv("LOCALAPPDATA", Path.home()))
    / "QualityAgentPublisher"
    / "config.json"
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def publish(package: Path, version: str, destination: Path, notes: str, required: bool):
    if not package.is_file():
        raise FileNotFoundError(f"找不到新版程序包：{package}")
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / package.name
    shutil.copy2(package, target)
    latest = {
        "version": version,
        "package": target.name,
        "sha256": sha256(target),
        "required": required,
        "notes": notes,
    }
    temp = destination / "latest.json.tmp"
    temp.write_text(json.dumps(latest, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, destination / "latest.json")
    (destination / f"更新说明-{version}.txt").write_text(notes or "本次版本未填写更新说明。", encoding="utf-8")
    return target


def load_last_destination() -> str:
    try:
        data = json.loads(PUBLISHER_CONFIG.read_text(encoding="utf-8"))
        return str(data.get("destination", "")).strip()
    except (OSError, json.JSONDecodeError, AttributeError):
        return ""


def save_last_destination(destination: str):
    PUBLISHER_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    PUBLISHER_CONFIG.write_text(
        json.dumps({"destination": destination}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def gui_main():
    try:
        package = win_dialogs.choose_file("选择新版程序ZIP")
        if not package:
            return
        package_path = Path(package)
        match = re.fullmatch(r"质检Agent-(.+)", package_path.stem)
        suggested_version = match.group(1) if match else ""
        version = win_dialogs.input_text(
            "版本号",
            "版本号已根据ZIP文件名自动填写，请确认后继续。",
            suggested_version,
        )
        if not version:
            return
        destination = win_dialogs.input_text(
            "共享发布文件夹",
            "请粘贴公司共享发布文件夹的完整地址。\n\n"
            "例如：\\\\公司文件服务器\\共享资料\\质检Agent发布\n\n"
            "程序会记住本次地址，下次发布时自动带出。",
            load_last_destination(),
        )
        destination = destination.strip().strip('"')
        if not destination:
            win_dialogs.warning("未发布", "没有选择共享发布文件夹，本次未发布任何文件。")
            return
        save_last_destination(destination)
        notes = win_dialogs.input_text("更新说明", "请简要填写本次更新内容") or ""
        required = win_dialogs.ask_yes_no("强制更新", "是否要求测试人员必须更新后才能启动？")
        target = publish(package_path, version.strip(), Path(destination), notes, required)
        win_dialogs.info("发布成功", f"新版已发布到：\n{target}")
    except Exception as exc:
        win_dialogs.error("发布失败", str(exc))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--package")
    parser.add_argument("--version")
    parser.add_argument("--destination")
    parser.add_argument("--notes", default="")
    parser.add_argument("--required", action="store_true")
    args = parser.parse_args()
    if args.package and args.version and args.destination:
        publish(Path(args.package), args.version, Path(args.destination), args.notes, args.required)
    else:
        gui_main()


if __name__ == "__main__":
    main()
