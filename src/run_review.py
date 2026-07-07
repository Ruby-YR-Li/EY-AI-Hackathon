"""Run the expense review assistant and export a traceable review package."""

from __future__ import annotations

import argparse
from pathlib import Path

from expense_review_engine import DEFAULT_FILES, review_files
from export_review_package import OUTPUT_DIR, export_review_package
from llm_assistant import DEFAULT_DEEPSEEK_BASE_URL, DEFAULT_DEEPSEEK_MODEL, LLMConfig


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="费用底稿 AI Review 助手")
    parser.add_argument(
        "--workpaper",
        type=Path,
        action="append",
        dest="workpapers",
        help="待 Review 费用底稿，可重复传入；支持主底稿和 TOD 底稿自动识别。",
    )
    parser.add_argument("--main-workpaper", type=Path, default=DEFAULT_FILES["main_workpaper"])
    parser.add_argument("--tod-workpaper", type=Path, default=DEFAULT_FILES["tod_workpaper"])
    parser.add_argument("--program-doc", type=Path, default=DEFAULT_FILES["program_doc"])
    parser.add_argument("--sop-workbook", type=Path, default=DEFAULT_FILES["sop_workbook"])
    parser.add_argument("--enable-llm", action="store_true", help="启用 DeepSeek 辅助判断。")
    parser.add_argument("--llm-api-key", default="", help="DeepSeek API Key；也可使用 DEEPSEEK_API_KEY 环境变量。")
    parser.add_argument("--llm-base-url", default=DEFAULT_DEEPSEEK_BASE_URL)
    parser.add_argument("--llm-model", default=DEFAULT_DEEPSEEK_MODEL)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    files = {
        "workpapers": args.workpapers,
        "main_workpaper": args.main_workpaper,
        "tod_workpaper": args.tod_workpaper,
        "program_doc": args.program_doc,
        "sop_workbook": args.sop_workbook,
    }
    llm_config = LLMConfig(
        enabled=args.enable_llm,
        api_key=args.llm_api_key,
        base_url=args.llm_base_url,
        model=args.llm_model,
    )
    summary, notes = review_files(files, llm_config=llm_config)
    outputs = export_review_package(notes, summary, files=files, output_dir=args.output_dir)

    print("Expense Review Assistant completed.")
    print(f"Review Notes: {summary.total_notes} | High: {summary.high_count} | Medium: {summary.medium_count}")
    for label, path in outputs.items():
        print(f"{label}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
