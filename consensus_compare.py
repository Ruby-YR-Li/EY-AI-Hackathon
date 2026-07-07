from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any


KEYS = [
    "false_positives",
    "false_negatives",
    "severity_disagreements",
    "category_disagreements",
    "prompt_issues",
    "script_issues",
    "recommended_fixes",
]


def normalize_text(value: Any) -> str:
    text = json.dumps(value, ensure_ascii=False, sort_keys=True) if not isinstance(value, str) else value
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"RN-\d+", "RN", text)
    return text.lower()


def canonical_items(report: dict, key: str) -> set[str]:
    items = report.get(key, [])
    if isinstance(items, dict):
        items = [items]
    if not isinstance(items, list):
        items = [items]
    return {normalize_text(item) for item in items if normalize_text(item)}


def compare_reports(a: dict, b: dict) -> dict:
    diffs = {}
    for key in KEYS:
        left = canonical_items(a, key)
        right = canonical_items(b, key)
        if left != right:
            diffs[key] = {
                "only_a": sorted(left - right),
                "only_b": sorted(right - left),
            }
    return {"consistent": not diffs, "diffs": diffs}


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: python consensus_compare.py agent_a.json agent_b.json")
        return 2
    a = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    b = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
    print(json.dumps(compare_reports(a, b), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
