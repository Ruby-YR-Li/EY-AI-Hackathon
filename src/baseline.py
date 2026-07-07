"""
项目mapping与本次质检基准值提取。

项目层保存mapping；每次质检任务上传最新版TE/SAD/A3表后，按mapping提取本次基准值。
"""

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Optional

import pandas as pd
from openpyxl.utils import column_index_from_string
from .app_paths import PROJECT_CONFIG_FILE as USER_PROJECT_CONFIG_FILE


PROJECT_CONFIG_FILE = str(USER_PROJECT_CONFIG_FILE)
PROJECT_CONFIG_DIR = os.path.dirname(PROJECT_CONFIG_FILE)


# 费用科目A3科目映射（销售费用&管理费用）
EXPENSE_A3_SUBJECTS = {
    "sales_expenses": ["销售费用"],
    "management_expenses": ["管理费用"],
}


def ensure_project_store() -> None:
    os.makedirs(PROJECT_CONFIG_DIR, exist_ok=True)
    if not os.path.exists(PROJECT_CONFIG_FILE):
        with open(PROJECT_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)


def load_projects() -> Dict[str, dict]:
    ensure_project_store()
    with open(PROJECT_CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_project(
    project_code: str,
    mapping: Optional[dict] = None,
    basic_info: Optional[dict] = None,
    materiality: Optional[dict] = None,
    subject_assertions: Optional[dict] = None,
) -> dict:
    ensure_project_store()
    project_code = project_code.strip()
    if not project_code:
        raise ValueError("项目代码不能为空")

    projects = load_projects()
    existing = projects.get(project_code, {})
    merged_basic_info = existing.get("basic_info", {}).copy()
    if basic_info is not None:
        for key, value in basic_info.items():
            if value not in (None, ""):
                merged_basic_info[key] = value
    merged_basic_info.pop("engagement_name", None)
    merged_basic_info.pop("project_name", None)
    record = {
        "project_code": project_code,
        "basic_info": merged_basic_info,
        "materiality": materiality if materiality is not None else existing.get("materiality", {}),
        "subject_assertions": (
            subject_assertions
            if subject_assertions is not None
            else existing.get("subject_assertions", {})
        ),
        "mapping": mapping if mapping not in (None, {}) else existing.get("mapping", {}),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    if project_code in projects and "created_at" in projects[project_code]:
        record["created_at"] = projects[project_code]["created_at"]
    else:
        record["created_at"] = record["updated_at"]
    projects[project_code] = record
    with open(PROJECT_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)
    return record


def get_project(project_code: str) -> Optional[dict]:
    return load_projects().get(project_code.strip())


def col_to_index(col: Any) -> int:
    if isinstance(col, int):
        return col - 1 if col > 0 else col
    text = str(col).strip()
    if text.isdigit():
        value = int(text)
        return value - 1 if value > 0 else value
    return column_index_from_string(text.upper()) - 1


def parse_number(value: Any) -> Optional[float]:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


@dataclass
class ExtractedValue:
    value: Optional[float]
    sheet: str = ""
    row: Optional[int] = None
    col: Optional[int] = None
    raw: Any = None

    def to_dict(self) -> dict:
        return {
            "value": self.value,
            "source": {
                "sheet": self.sheet,
                "row": self.row,
                "col": self.col,
                "raw": "" if self.raw is None else str(self.raw),
            },
        }


class BaselineExtractor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.xls = pd.ExcelFile(file_path)
        self.sheet_names = self.xls.sheet_names
        self._cache: Dict[str, pd.DataFrame] = {}

    def _sheet_name(self, sheet: Optional[str]) -> str:
        if not sheet or sheet == "*":
            return self.sheet_names[0]
        for actual in self.sheet_names:
            if actual.strip() == str(sheet).strip():
                return actual
        for actual in self.sheet_names:
            if str(sheet).strip() in actual.strip() or actual.strip() in str(sheet).strip():
                return actual
        raise ValueError(f"未找到sheet: {sheet}")

    def _df(self, sheet: str) -> pd.DataFrame:
        actual = self._sheet_name(sheet)
        if actual not in self._cache:
            self._cache[actual] = pd.read_excel(self.xls, sheet_name=actual, header=None)
        return self._cache[actual]

    def extract_value(self, mapping: Optional[dict]) -> ExtractedValue:
        if not mapping:
            return ExtractedValue(None)
        method = mapping.get("method", "cell")
        sheet = self._sheet_name(mapping.get("sheet"))
        df = self._df(sheet)

        if method == "cell":
            cell = str(mapping.get("cell", "")).strip()
            match = re.fullmatch(r"([A-Za-z]+)(\d+)", cell)
            if not match:
                raise ValueError(f"无效单元格定位: {cell}")
            col = col_to_index(match.group(1))
            row = int(match.group(2)) - 1
            raw = df.iloc[row, col]
            return ExtractedValue(parse_number(raw), sheet, row + 1, col + 1, raw)

        if method == "row_col_keyword":
            row_keyword = str(mapping.get("row_keyword", "")).strip()
            col_keyword = str(mapping.get("col_keyword", "")).strip()
            row_idx = self._find_row_by_keyword(df, row_keyword)
            col_idx = self._find_col_by_keyword(df, col_keyword)
            raw = df.iloc[row_idx, col_idx]
            return ExtractedValue(parse_number(raw), sheet, row_idx + 1, col_idx + 1, raw)

        raise ValueError(f"不支持的mapping方法: {method}")

    def _find_row_by_keyword(self, df: pd.DataFrame, keyword: str) -> int:
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                value = df.iloc[row_idx, col_idx]
                if value is not None and pd.notna(value) and keyword in str(value):
                    return row_idx
        raise ValueError(f"未找到行关键词: {keyword}")

    def _find_col_by_keyword(self, df: pd.DataFrame, keyword: str) -> int:
        for row_idx in range(min(30, len(df))):
            for col_idx in range(len(df.columns)):
                value = df.iloc[row_idx, col_idx]
                if value is not None and pd.notna(value) and keyword in str(value):
                    return col_idx
        raise ValueError(f"未找到列关键词: {keyword}")

    def extract_a3(self, a3_mapping: Optional[dict], subject: str) -> dict:
        if not a3_mapping:
            return {}
        if subject not in ("fixed_assets", "management_sales_expenses"):
            raise ValueError("当前仅支持固定资产或费用科目")

        sheet = self._sheet_name(a3_mapping.get("sheet"))
        df = self._df(sheet)
        subject_col = col_to_index(a3_mapping["subject_name_col"])
        book_col = col_to_index(a3_mapping["book_amount_col"])
        audited_col = col_to_index(a3_mapping["audited_amount_col"])
        prior_col = col_to_index(a3_mapping["prior_audited_amount_col"])

        a3_subjects = EXPENSE_A3_SUBJECTS if subject == "management_sales_expenses" else FIXED_ASSET_A3_SUBJECTS
        extracted = {}
        for key, keywords in a3_subjects.items():
            row_idx = self._find_subject_row(df, subject_col, keywords)
            if row_idx is None:
                extracted[key] = {"missing": True, "keywords": keywords}
                continue

            extracted[key] = {
                "book_amount": self._value_at(df, sheet, row_idx, book_col).to_dict(),
                "audited_amount": self._value_at(df, sheet, row_idx, audited_col).to_dict(),
                "prior_audited_amount": self._value_at(df, sheet, row_idx, prior_col).to_dict(),
                "subject_name": str(df.iloc[row_idx, subject_col]),
            }
        return extracted

    def _find_subject_row(self, df: pd.DataFrame, subject_col: int, keywords: list) -> Optional[int]:
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, subject_col] if subject_col < len(df.columns) else None
            if value is None or pd.isna(value):
                continue
            text = str(value).strip()
            if any(keyword in text for keyword in keywords):
                return row_idx
        return None

    def _value_at(self, df: pd.DataFrame, sheet: str, row_idx: int, col_idx: int) -> ExtractedValue:
        raw = df.iloc[row_idx, col_idx]
        return ExtractedValue(parse_number(raw), sheet, row_idx + 1, col_idx + 1, raw)


def extract_task_baseline(file_path: str, mapping: dict, subject: str,
                          include_te_sad: bool = True, include_a3: bool = True) -> dict:
    extractor = BaselineExtractor(file_path)
    baseline = {}
    if include_te_sad:
        baseline["te"] = extractor.extract_value(mapping.get("te_mapping")).to_dict()
        baseline["sad"] = extractor.extract_value(mapping.get("sad_mapping")).to_dict()
    if include_a3:
        baseline["a3"] = extractor.extract_a3(mapping.get("a3_mapping"), subject)
    return baseline
