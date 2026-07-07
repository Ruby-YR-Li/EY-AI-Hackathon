"""Zero-dependency local upload page for the expense review assistant."""

from __future__ import annotations

import cgi
import html
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

from expense_review_engine import DEFAULT_FILES, review_files
from export_review_package import OUTPUT_DIR, export_review_package
from llm_assistant import DEFAULT_DEEPSEEK_BASE_URL, DEFAULT_DEEPSEEK_MODEL, LLMConfig


HOST = "127.0.0.1"
PORT = 8765
UPLOAD_DIR = Path(tempfile.gettempdir()) / "expense_review_uploads"


def page(message: str = "") -> bytes:
    msg = f"<p class='msg'>{message}</p>" if message else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>费用底稿 AI Review 助手</title>
  <style>
    body {{ margin:0; background:#f5f5f5; font-family:"Segoe UI","Microsoft YaHei",Arial,sans-serif; color:#111; }}
    header {{ background:#111; color:white; border-left:10px solid #ffe600; padding:22px 30px; }}
    main {{ max-width:980px; margin:24px auto; background:white; border:1px solid #ddd; border-radius:8px; padding:22px; }}
    label {{ display:block; margin-top:16px; font-weight:700; }}
    input {{ margin-top:6px; width:100%; padding:8px; border:1px solid #ccc; border-radius:4px; }}
    input[type="checkbox"] {{ width:auto; margin-right:8px; }}
    button {{ margin-top:20px; background:#111; color:#fff; border:0; padding:11px 18px; border-radius:4px; cursor:pointer; }}
    .msg {{ background:#fffbea; border-left:5px solid #ffe600; padding:10px 12px; }}
    .hint {{ color:#667085; font-size:14px; }}
  </style>
</head>
<body>
  <header>
    <h1>费用底稿 AI Review 助手</h1>
    <p>上传测试底稿，生成 Review Report、Review Notes Excel 和标注底稿副本。</p>
  </header>
  <main>
    {msg}
    <form method="post" enctype="multipart/form-data">
      <p class="hint">SOP、审计程序要求和标准底稿已作为后台质检资料内置，不需要上传。</p>
      <label>待 Review 费用底稿（必填，可多选 xlsx）</label>
      <input type="file" name="workpapers" accept=".xlsx,.xlsm" multiple required>
      <p class="hint">可同时选择主费用底稿和 TOD 抽样底稿，系统会自动识别为同一套底稿。</p>
      <label>序时账（可选 xlsx）</label>
      <input type="file" name="general_ledger" accept=".xlsx,.xlsm">
      <label>科目余额表（可选 xlsx）</label>
      <input type="file" name="trial_balance" accept=".xlsx,.xlsm">
      <label><input type="checkbox" name="llm_enabled" checked>启用 DeepSeek 辅助判断</label>
      <label>DeepSeek API Key（本次使用，不保存）</label>
      <input type="password" name="llm_api_key" placeholder="sk-...">
      <label>DeepSeek Base URL</label>
      <input type="text" name="llm_base_url" value="{html.escape(DEFAULT_DEEPSEEK_BASE_URL)}">
      <label>模型</label>
      <input type="text" name="llm_model" value="{html.escape(DEFAULT_DEEPSEEK_MODEL)}">
      <button type="submit">开始 Review</button>
    </form>
  </main>
</body>
</html>""".encode("utf-8")


def save_field(form: cgi.FieldStorage, name: str, default_path: Path) -> Path:
    field = form[name] if name in form else None
    if field is None or not getattr(field, "filename", ""):
        return default_path
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    suffix = Path(field.filename).suffix or default_path.suffix
    target = UPLOAD_DIR / f"{name}{suffix}"
    with target.open("wb") as fh:
        fh.write(field.file.read())
    return target


def save_optional_field(form: cgi.FieldStorage, name: str) -> Path | None:
    field = form[name] if name in form else None
    if isinstance(field, list):
        field = next((item for item in field if getattr(item, "filename", "")), None)
    if field is None or not getattr(field, "filename", ""):
        return None
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    filename = Path(field.filename).name
    target = UPLOAD_DIR / f"{name}_{filename}"
    with target.open("wb") as fh:
        fh.write(field.file.read())
    return target


def save_workpaper_fields(form: cgi.FieldStorage) -> list[Path]:
    fields = form["workpapers"] if "workpapers" in form else []
    if not isinstance(fields, list):
        fields = [fields]
    saved: list[Path] = []
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    for index, field in enumerate(fields, start=1):
        if not getattr(field, "filename", ""):
            continue
        filename = Path(field.filename).name
        target = UPLOAD_DIR / f"workpaper_{index}_{filename}"
        with target.open("wb") as fh:
            fh.write(field.file.read())
        saved.append(target)
    return saved


def form_value(form: cgi.FieldStorage, name: str, default: str = "") -> str:
    field = form[name] if name in form else None
    if isinstance(field, list):
        field = field[0] if field else None
    if field is None:
        return default
    return str(getattr(field, "value", default) or default).strip()


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path.startswith("/outputs/"):
            rel = unquote(self.path.removeprefix("/outputs/"))
            target = (OUTPUT_DIR / rel).resolve()
            if OUTPUT_DIR.resolve() not in target.parents and target != OUTPUT_DIR.resolve():
                self.send_error(403)
                return
            if not target.is_file():
                self.send_error(404)
                return
            self.send_response(200)
            if target.suffix.lower() == ".html":
                self.send_header("Content-Type", "text/html; charset=utf-8")
            else:
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
            self.end_headers()
            self.wfile.write(target.read_bytes())
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(page())

    def do_POST(self) -> None:
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )
            workpapers = save_workpaper_fields(form)
            if not workpapers:
                raise ValueError("请至少上传一份待 Review 费用底稿。")
            files = {
                "workpapers": workpapers,
                "trial_balance": save_optional_field(form, "trial_balance"),
                "general_ledger": save_optional_field(form, "general_ledger"),
                "program_doc": DEFAULT_FILES["program_doc"],
                "sop_workbook": DEFAULT_FILES["sop_workbook"],
            }
            llm_config = LLMConfig(
                enabled="llm_enabled" in form,
                api_key=form_value(form, "llm_api_key"),
                base_url=form_value(form, "llm_base_url", DEFAULT_DEEPSEEK_BASE_URL),
                model=form_value(form, "llm_model", DEFAULT_DEEPSEEK_MODEL),
            )
            summary, notes = review_files(files, llm_config=llm_config)
            outputs = export_review_package(notes, summary, files=files, output_dir=OUTPUT_DIR)
            links = "<br>".join(
                f'<a href="/outputs/{html.escape(path.name)}">{html.escape(label)}: {html.escape(path.name)}</a>'
                for label, path in outputs.items()
            )
            llm_status = "DeepSeek 已启用" if llm_config.is_available else "DeepSeek 未配置，已使用规则引擎完成"
            message = f"Review 完成：{summary.total_notes} 条 Review Notes，其中 High {summary.high_count} 条。{llm_status}<br>{links}"
        except Exception as exc:
            message = f"Review 失败：{exc}"

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(page(message))


def main() -> int:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Expense Review Assistant running at http://{HOST}:{PORT}")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
