"""质检Agent固定启动器：首次配置共享路径、自动更新、启动本地网页。"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser
import zipfile
from pathlib import Path
import win_dialogs


APP_HOME = Path(os.getenv("LOCALAPPDATA", Path.home())) / "QualityAgent"
CONFIG_DIR = APP_HOME / "config"
VERSIONS_DIR = APP_HOME / "versions"
UPDATE_DIR = APP_HOME / "updates"
LOG_DIR = APP_HOME / "logs"
LAUNCHER_CONFIG = CONFIG_DIR / "launcher_config.json"
CURRENT_FILE = APP_HOME / "current.json"
SERVER_URL = "http://127.0.0.1:8000"


def ensure_dirs():
    for path in (APP_HOME, CONFIG_DIR, VERSIONS_DIR, UPDATE_DIR, LOG_DIR):
        path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default=None):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        pass
    return {} if default is None else default


def save_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def log(message: str):
    ensure_dirs()
    with (LOG_DIR / "launcher.log").open("a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")


def version_key(version: str):
    parts = []
    for token in version.replace("-", ".").split("."):
        parts.append(int(token) if token.isdigit() else token.lower())
    return tuple(parts)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def port_open() -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.3)
        return sock.connect_ex(("127.0.0.1", 8000)) == 0


def stop_running_server():
    """关闭占用8000端口的旧服务，确保启动的是本机当前版本。"""
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    completed = subprocess.run(
        ["netstat", "-ano", "-p", "tcp"],
        check=False,
        capture_output=True,
        text=True,
        creationflags=creationflags,
    )
    pids = set()
    for line in (completed.stdout or "").splitlines():
        parts = line.split()
        if len(parts) < 5 or parts[0].upper() != "TCP":
            continue
        if parts[1].endswith(":8000") and parts[3].upper() == "LISTENING":
            pids.add(parts[4])
    for pid in pids:
        subprocess.run(
            ["taskkill", "/F", "/PID", pid],
            check=False,
            capture_output=True,
            creationflags=creationflags,
        )
        log(f"启动前关闭旧服务 pid={pid}")
    for _ in range(30):
        if not port_open():
            return
        time.sleep(0.2)
    if pids:
        raise RuntimeError("旧版服务未能关闭。请重启电脑后再打开质检Agent。")


def prompt_update_source() -> str:
    config = load_json(LAUNCHER_CONFIG)
    current = str(config.get("update_source", "")).strip()
    if current:
        return current
    value = win_dialogs.input_text(
        "首次启动设置",
        "请输入公司共享版本文件夹的完整地址。\n\n"
        "例如：\\\\公司文件服务器\\共享资料\\质检Agent发布\n\n"
        "如果暂时没有地址，可以点“取消”，先使用当前本地版本。",
    )
    if value:
        value = os.path.expandvars(value.strip())
        config["update_source"] = value
        save_json(LAUNCHER_CONFIG, config)
        return value
    return ""


def choose_update_source() -> str:
    selected = win_dialogs.choose_folder("选择公司共享版本文件夹")
    if selected:
        config = load_json(LAUNCHER_CONFIG)
        config["update_source"] = selected
        save_json(LAUNCHER_CONFIG, config)
    return selected


def safe_extract(zip_path: Path, target: Path):
    target_resolved = target.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            destination = (target / member.filename).resolve()
            if destination != target_resolved and target_resolved not in destination.parents:
                raise ValueError(f"压缩包包含不安全路径: {member.filename}")
        archive.extractall(target)


def install_update(source: Path, latest: dict) -> str:
    version = str(latest["version"]).strip()
    package_name = str(latest["package"]).strip()
    package_source = source / package_name
    if not package_source.is_file():
        raise FileNotFoundError(f"共享文件夹中未找到新版程序包：{package_name}")

    local_zip = UPDATE_DIR / package_name
    shutil.copy2(package_source, local_zip)
    expected_hash = str(latest.get("sha256", "")).strip().lower()
    if expected_hash and sha256(local_zip).lower() != expected_hash:
        raise ValueError("新版程序包校验失败，请联系发布人员重新发布")

    version_dir = VERSIONS_DIR / version
    temp_dir = VERSIONS_DIR / f".installing-{version}"
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(parents=True)
    safe_extract(local_zip, temp_dir)
    app_root = temp_dir
    children = [p for p in temp_dir.iterdir()]
    if len(children) == 1 and children[0].is_dir():
        app_root = children[0]
    if version_dir.exists():
        shutil.rmtree(version_dir)
    if app_root == temp_dir:
        os.replace(temp_dir, version_dir)
    else:
        os.replace(app_root, version_dir)
        shutil.rmtree(temp_dir, ignore_errors=True)
    save_json(CURRENT_FILE, {"version": version, "path": str(version_dir)})
    log(f"更新完成 version={version}")
    return version


def check_update(update_source: str):
    if not update_source:
        return False
    source = Path(update_source)
    if not source.is_dir():
        choice = win_dialogs.ask_yes_no_cancel(
            "无法访问共享版本文件夹",
            "当前无法访问已配置的共享版本文件夹。\n\n"
            "选择“是”：重新选择文件夹\n"
            "选择“否”：继续使用本地版本\n"
            "选择“取消”：退出",
        )
        if choice is None:
            raise SystemExit(0)
        if choice:
            new_source = choose_update_source()
            if new_source:
                return check_update(new_source)
        return False

    latest_file = source / "latest.json"
    if not latest_file.exists():
        log(f"共享目录暂无latest.json: {source}")
        return False
    latest = load_json(latest_file)
    latest_version = str(latest.get("version", "")).strip()
    if not latest_version or not latest.get("package"):
        raise ValueError("共享版本文件latest.json内容不完整")
    current_version = str(load_json(CURRENT_FILE).get("version", "0"))
    if version_key(latest_version) <= version_key(current_version):
        return False
    mandatory = bool(latest.get("required", False))
    if not mandatory:
        ok = win_dialogs.ask_yes_no(
            "发现新版",
            f"发现质检Agent新版本 {latest_version}。\n\n是否现在自动更新？",
        )
        if not ok:
            return False
    install_update(source, latest)
    win_dialogs.info("更新完成", f"已更新到版本 {latest_version}。")
    return True


def current_server_command():
    current = load_json(CURRENT_FILE)
    version_dir = Path(str(current.get("path", "")))
    candidates = [
        version_dir / "质检Agent服务.exe",
        version_dir / "质检Agent服务" / "质检Agent服务.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return [str(candidate)]
    # 开发环境回退：直接运行当前工作区源码。
    source_root = Path(__file__).resolve().parent
    run_server = source_root / "run_server.py"
    if run_server.exists():
        return [sys.executable, str(run_server)]
    raise FileNotFoundError("本机没有可用的质检Agent版本，请重新运行首次安装包")


def start_server():
    if port_open():
        webbrowser.open(SERVER_URL)
        return
    command = current_server_command()
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        command,
        cwd=str(Path(command[-1]).resolve().parent),
        creationflags=creationflags,
    )
    for _ in range(60):
        try:
            with urllib.request.urlopen(SERVER_URL, timeout=1) as response:
                if response.status == 200:
                    webbrowser.open(SERVER_URL)
                    return
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("质检网页启动超时，请查看本机QualityAgent\\logs目录中的日志")


def main():
    ensure_dirs()
    try:
        update_source = prompt_update_source()
        check_update(update_source)
        # 即使本次没有下载更新，也可能有覆盖安装前遗留的旧版服务。
        # 每次从桌面启动都重启后台，保证网页对应current.json中的当前版本。
        stop_running_server()
        start_server()
    except SystemExit:
        return
    except Exception as exc:
        log(f"启动失败: {exc}")
        win_dialogs.error("质检Agent启动失败", str(exc))


if __name__ == "__main__":
    main()
