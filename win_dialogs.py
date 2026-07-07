"""不依赖Tkinter的Windows原生对话框。"""

from __future__ import annotations

import ctypes
import subprocess


MB_OK = 0x00000000
MB_ICONINFORMATION = 0x00000040
MB_ICONWARNING = 0x00000030
MB_ICONERROR = 0x00000010
MB_YESNO = 0x00000004
MB_YESNOCANCEL = 0x00000003

IDYES = 6
IDNO = 7
IDCANCEL = 2


def _ps_quote(value: str) -> str:
    return str(value).replace("'", "''")


def _run_powershell(script: str) -> str:
    completed = subprocess.run(
        ["powershell", "-STA", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    # 部分企业内网电脑或PyInstaller窗口程序中，PowerShell即使正常结束，
    # stdout也可能返回None。此时应视为用户未选择，而不是让工具崩溃。
    return (completed.stdout or "").strip()


def info(title: str, text: str):
    ctypes.windll.user32.MessageBoxW(None, text, title, MB_OK | MB_ICONINFORMATION)


def warning(title: str, text: str):
    ctypes.windll.user32.MessageBoxW(None, text, title, MB_OK | MB_ICONWARNING)


def error(title: str, text: str):
    ctypes.windll.user32.MessageBoxW(None, text, title, MB_OK | MB_ICONERROR)


def ask_yes_no(title: str, text: str) -> bool:
    result = ctypes.windll.user32.MessageBoxW(None, text, title, MB_YESNO | MB_ICONINFORMATION)
    return result == IDYES


def ask_yes_no_cancel(title: str, text: str):
    result = ctypes.windll.user32.MessageBoxW(None, text, title, MB_YESNOCANCEL | MB_ICONWARNING)
    if result == IDYES:
        return True
    if result == IDNO:
        return False
    return None


def input_text(title: str, prompt: str, default: str = "") -> str:
    script = (
        "Add-Type -AssemblyName Microsoft.VisualBasic;"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        f"$v=[Microsoft.VisualBasic.Interaction]::InputBox('{_ps_quote(prompt)}',"
        f"'{_ps_quote(title)}','{_ps_quote(default)}');"
        "[Console]::Write($v)"
    )
    return _run_powershell(script)


def choose_folder(title: str) -> str:
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
        f"$d.Description='{_ps_quote(title)}';$d.ShowNewFolderButton=$false;"
        "if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)"
        "{[Console]::Write($d.SelectedPath)}"
    )
    return _run_powershell(script)


def choose_file(title: str, filter_text: str = "ZIP程序包 (*.zip)|*.zip") -> str:
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "$d=New-Object System.Windows.Forms.OpenFileDialog;"
        f"$d.Title='{_ps_quote(title)}';$d.Filter='{_ps_quote(filter_text)}';"
        "if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)"
        "{[Console]::Write($d.FileName)}"
    )
    return _run_powershell(script)
