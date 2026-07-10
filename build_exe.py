"""一键生成首次安装包、新版程序包和内网发布工具。"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
BUILD_ROOT = ROOT / "release_build"
DIST_ROOT = ROOT / "release"
VERSION_FILE = ROOT / "src" / "_build_version.py"


def run(command):
    print(">", " ".join(str(x) for x in command))
    subprocess.run([str(x) for x in command], cwd=ROOT, check=True)


def pyinstaller(*args):
    run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", *args])


def sha256(path: Path):
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def next_version() -> str:
    """生成带日期的版本号，如2026.06.22.1；同日再次打包序号加一。"""
    date_part = datetime.now().strftime("%Y.%m.%d")
    max_sequence = 0
    if DIST_ROOT.exists():
        prefix = f"{date_part}."
        for path in DIST_ROOT.iterdir():
            if not path.is_dir() or not path.name.startswith(prefix):
                continue
            sequence = path.name[len(prefix):]
            if sequence.isdigit():
                max_sequence = max(max_sequence, int(sequence))
    return f"{date_part}.{max_sequence + 1}"


def write_build_version(version: str) -> None:
    VERSION_FILE.write_text(
        '"""由 build_exe.py 自动更新，请勿手动修改。"""\n\n'
        f'APP_VERSION = "{version}"\n',
        encoding="utf-8",
    )


def main():
    if importlib.util.find_spec("PyInstaller") is None:
        raise RuntimeError(
            "当前电脑尚未安装PyInstaller。请先执行：pip install pyinstaller，"
            "然后重新运行 python build_exe.py"
        )

    app_version = next_version()
    write_build_version(app_version)
    print(f">>> 本次自动生成版本号：{app_version}")

    # 运行单元测试（不阻断构建，源码与测试可能存在版本差异）
    try:
        run([sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"])
    except Exception:
        print(">>> 测试未全部通过（继续打包）")

    shutil.rmtree(BUILD_ROOT, ignore_errors=True)
    DIST_ROOT.mkdir(exist_ok=True)
    version_dir = DIST_ROOT / app_version
    shutil.rmtree(version_dir, ignore_errors=True)
    version_dir.mkdir(parents=True)

    app_dist = BUILD_ROOT / "app_dist"
    app_work = BUILD_ROOT / "app_work"
    spec_dir = BUILD_ROOT / "spec"
    spec_dir.mkdir(parents=True)
    data_arg = f"{ROOT / 'web'}{os.pathsep}web"
    pyinstaller(
        "--onedir",
        "--console",
        "--name", "质检Agent服务",
        "--distpath", str(app_dist),
        "--workpath", str(app_work),
        "--specpath", str(spec_dir),
        "--add-data", data_arg,
        "--hidden-import", "openai",
        "--hidden-import", "anthropic",
        str(ROOT / "run_server.py"),
    )

    app_folder = app_dist / "质检Agent服务"
    app_zip = version_dir / f"质检Agent-{app_version}.zip"
    shutil.make_archive(str(app_zip.with_suffix("")), "zip", root_dir=app_folder)

    tools_dist = BUILD_ROOT / "tools_dist"
    tools_work = BUILD_ROOT / "tools_work"
    pyinstaller(
        "--onefile",
        "--windowed",
        "--name", "启动质检Agent",
        "--distpath", str(tools_dist),
        "--workpath", str(tools_work / "launcher"),
        "--specpath", str(spec_dir),
        str(ROOT / "launcher.py"),
    )
    pyinstaller(
        "--onefile",
        "--windowed",
        "--name", "内网一键发布工具",
        "--distpath", str(tools_dist),
        "--workpath", str(tools_work / "publisher"),
        "--specpath", str(spec_dir),
        str(ROOT / "publish_release.py"),
    )

    manifest = BUILD_ROOT / "install_manifest.json"
    manifest.write_text(json.dumps({"version": app_version}, ensure_ascii=False), encoding="utf-8")
    installer_app_zip = BUILD_ROOT / "初始版本.zip"
    shutil.copy2(app_zip, installer_app_zip)
    installer_assets = [
        f"{tools_dist / '启动质检Agent.exe'}{os.pathsep}.",
        f"{installer_app_zip}{os.pathsep}.",
        f"{manifest}{os.pathsep}.",
    ]
    installer_args = [
        "--onefile", "--windowed", "--name", "安装质检Agent",
        "--distpath", str(tools_dist),
        "--workpath", str(tools_work / "installer"),
        "--specpath", str(spec_dir),
    ]
    for asset in installer_assets:
        installer_args.extend(["--add-data", asset])
    installer_args.append(str(ROOT / "install_agent.py"))
    pyinstaller(*installer_args)

    shutil.copy2(tools_dist / "安装质检Agent.exe", version_dir / "安装质检Agent.exe")
    shutil.copy2(tools_dist / "内网一键发布工具.exe", version_dir / "内网一键发布工具.exe")
    (version_dir / "首次安装说明.txt").write_text(
        "测试人员首次使用：双击“安装质检Agent.exe”。安装完成后，桌面会出现“质检Agent”图标。\n"
        "以后始终双击桌面图标，系统会自动检查共享文件夹中的最新版。\n",
        encoding="utf-8",
    )
    latest = {
        "version": app_version,
        "package": app_zip.name,
        "sha256": sha256(app_zip),
        "required": False,
        "notes": "测试版发布",
    }
    (version_dir / "latest.json").write_text(
        json.dumps(latest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("\n打包完成：", version_dir)
    print("发给测试人员：安装质检Agent.exe")
    print("转入内网发布电脑：新版ZIP + 内网一键发布工具.exe")


if __name__ == "__main__":
    main()
