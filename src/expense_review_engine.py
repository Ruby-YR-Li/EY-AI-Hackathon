"""Expense workpaper review rules for the hackathon MVP.

The engine is intentionally case-driven: it covers the supplied VC&VD workpaper
set first, while keeping rules small and auditable for later expansion.
"""

from __future__ import annotations

import csv
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, Sequence

import openpyxl

from llm_assistant import LLMConfig

try:
    from docx import Document
except Exception:  # pragma: no cover - optional dependency in CLI smoke checks.
    Document = None


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "资料库" / "AI底稿Review助手-输入材料"
SOP_DIR = ROOT / "资料库" / "SOP"

DEFAULT_FILES = {
    "main_workpaper": DATA_DIR / "U_exp SWP VC&VD 20251231 .xlsx",
    "tod_workpaper": DATA_DIR / "TOD SWP U_Exp 20251231 .xlsx",
    "program_doc": DATA_DIR / "审计程序要求.docx",
    "sop_workbook": SOP_DIR / "FY26_SOP U_exp SWP VC&VD.xlsx",
}

LEGAL_KEYWORDS = ("法律", "律师", "诉讼", "咨询", "中介")
INTERNAL_SHEET_PREFIXES = ("Skywind", "DS_INTERNAL")


@dataclass
class ReviewNote:
    risk_level: str
    file: str
    sheet: str
    location: str
    issue_type: str
    review_note: str
    suggested_action: str
    source: str
    judge_method: str = "Rule"
    confidence: str = ""
    evidence_summary: str = ""


@dataclass
class ReviewSummary:
    total_notes: int
    high_count: int
    medium_count: int
    low_count: int
    sheets_impacted: int
    rules_run: int
    program_rows: int
    checklist_rows: int


@dataclass
class WorkpaperSelection:
    main_workpaper: Path
    tod_workpaper: Path | None
    uploaded_workpapers: list[Path]


def cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def is_blank(value: object) -> bool:
    return cell_text(value) == ""


def as_number(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def load_workbook(path: Path, *, data_only: bool = True):
    return openpyxl.load_workbook(path, data_only=data_only, read_only=True)


def _sheetnames(path: Path) -> list[str]:
    wb = load_workbook(path)
    return list(wb.sheetnames)


def classify_workpaper(path: Path) -> str:
    """Classify uploaded workpaper files by filename and sheet names."""
    name_text = path.name.lower()
    sheets = _sheetnames(path)
    sheet_text = " ".join(sheets).lower()
    main_markers = ("Uexp.00 Lead", "VC.00 销售费用BKD", "VD.00 管理费用BKD", "汇总")
    if any(marker in sheets for marker in main_markers):
        return "main"
    if any(token in name_text or token in sheet_text for token in ("tod", "抽样", "样本", "test of details", "skywind")):
        return "tod"
    return "unknown"


def select_workpapers(paths: dict[str, object]) -> WorkpaperSelection:
    uploaded = [Path(p) for p in paths.get("workpapers", []) or []]
    if not uploaded:
        main = Path(paths.get("main_workpaper") or DEFAULT_FILES["main_workpaper"])
        tod = paths.get("tod_workpaper")
        return WorkpaperSelection(main, Path(tod) if tod else None, [main] + ([Path(tod)] if tod else []))

    main_files: list[Path] = []
    tod_files: list[Path] = []
    for path in uploaded:
        file_type = classify_workpaper(path)
        if file_type == "main":
            main_files.append(path)
        elif file_type == "tod":
            tod_files.append(path)

    if len(main_files) > 1:
        names = ", ".join(p.name for p in main_files)
        raise ValueError(f"识别到多个主费用底稿：{names}。请仅上传同一套底稿中的一份主底稿。")
    if not main_files:
        raise ValueError("未识别到主费用底稿。请上传包含 Uexp.00 Lead / VC.00 / VD.00 / 汇总 等 sheet 的费用底稿。")

    return WorkpaperSelection(main_files[0], tod_files[0] if tod_files else None, uploaded)


def note(
    risk_level: str,
    file: Path,
    sheet: str,
    location: str,
    issue_type: str,
    review_note: str,
    suggested_action: str,
    source: str,
    judge_method: str = "Rule",
    confidence: str = "",
    evidence_summary: str = "",
) -> ReviewNote:
    return ReviewNote(
        risk_level=risk_level,
        file=file.name,
        sheet=sheet,
        location=location,
        issue_type=issue_type,
        review_note=review_note,
        suggested_action=suggested_action,
        source=source,
        judge_method=judge_method,
        confidence=confidence,
        evidence_summary=evidence_summary,
    )


def read_program_rows(path: Path) -> int:
    if not path.exists() or Document is None:
        return 0
    doc = Document(str(path))
    if not doc.tables:
        return 0
    return max(0, len(doc.tables[0].rows) - 1)


def read_review_checklist_rows(path: Path) -> int:
    if not path.exists():
        return 0
    wb = load_workbook(path)
    if "Review checklist" not in wb.sheetnames:
        return 0
    ws = wb["Review checklist"]
    count = 0
    for row in ws.iter_rows(min_row=1, values_only=True):
        text = " ".join(cell_text(v) for v in row)
        if "◆" in text or "检查" in text:
            count += 1
    return count


def review_lead_basics(wb, file: Path) -> list[ReviewNote]:
    sheet = "Uexp.00 Lead"
    if sheet not in wb.sheetnames:
        return [
            note(
                "High",
                file,
                sheet,
                "Sheet",
                "缺少关键底稿页",
                "主底稿缺少 Uexp.00 Lead，无法完成基础信息和 Lead 勾稽复核。",
                "确认使用的是否为销售费用/管理费用标准底稿，或补充 Lead 页。",
                "SOP Excel - Uexp.00 Lead",
            )
        ]

    ws = wb[sheet]
    notes: list[ReviewNote] = []
    required_cells = {
        "C2": "客户名称",
        "C3": "期末",
        "C4": "分析日期",
        "C5": "可容忍误差(TE)",
        "C6": "名义金额(SAD)",
        "C7": "适用会计准则",
        "C8": "记账本位币",
    }
    for coord, label in required_cells.items():
        if is_blank(ws[coord].value):
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    coord,
                    "Lead 基础信息缺失",
                    f"{label}未填写，底稿基础信息不完整。",
                    f"在 {coord} 补充{label}，并与项目基础资料或 Canvas 设置核对。",
                    "SOP - Lead sheet 易错点",
                )
            )

    accounting_standard = cell_text(ws["C7"].value).upper()
    functional_currency = cell_text(ws["C8"].value).upper()
    currency_values = {"CNY", "RMB", "人民币", "USD", "EUR", "HKD"}
    standard_values = {"IFRS", "CAS", "PRC GAAP", "企业会计准则", "中国企业会计准则"}
    if accounting_standard in currency_values:
        notes.append(
            note(
                "High",
                file,
                sheet,
                "C7",
                "基础信息疑似填列错误",
                f"适用会计准则填写为“{ws['C7'].value}”，疑似将会计准则填成货币。",
                "C7 应填写企业会计准则、IFRS 等适用准则；同时核对 C8 是否为记账本位币。",
                "SOP - Lead sheet 基础信息",
            )
        )
    if functional_currency in standard_values:
        notes.append(
            note(
                "High",
                file,
                sheet,
                "C8",
                "基础信息疑似填列错误",
                f"记账本位币填写为“{ws['C8'].value}”，疑似将记账本位币填成会计准则。",
                "C8 应填写 CNY、RMB、人民币等记账本位币；同时核对 C7 是否为适用会计准则。",
                "SOP - Lead sheet 基础信息",
            )
        )

    if as_number(ws["C5"].value) is not None and as_number(ws["C6"].value) is not None:
        if as_number(ws["C6"].value) > as_number(ws["C5"].value):
            notes.append(
                note(
                    "Medium",
                    file,
                    sheet,
                    "C5:C6",
                    "TE/SAD 关系提示",
                    "SAD 大于 TE，需确认项目重要性参数是否填列正确。",
                    "与 Canvas 或项目 PM/TE/SAD 底稿核对，确认是否需要更新。",
                    "SOP - PM/TE/SAD 易错点",
                )
            )

    return notes


def review_summary_execution(wb, file: Path) -> list[ReviewNote]:
    sheet = "汇总"
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    notes: list[ReviewNote] = []
    for row in range(2, ws.max_row + 1):
        procedure_page = cell_text(ws.cell(row, 6).value)
        execution = cell_text(ws.cell(row, 7).value)
        reason = cell_text(ws.cell(row, 8).value)
        if not procedure_page or procedure_page == "程序页":
            continue
        if not execution:
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    f"G{row}",
                    "程序执行状态缺失",
                    f"程序“{procedure_page}”未填写是否执行。",
                    "在执行列选择“是/否”；如不执行，在不执行原因列补充具体理由。",
                    "审计程序要求 + SOP 汇总页",
                )
            )
        if execution == "否" and not reason:
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    f"H{row}",
                    "拒绝执行理由缺失",
                    f"程序“{procedure_page}”选择不执行，但未填写不执行原因。",
                    "补充不执行的业务原因、审计判断和替代程序依据。",
                    "SOP 汇总页易错点",
                )
            )
    return notes


def find_special_expenses(wb, file: Path) -> list[ReviewNote]:
    notes: list[ReviewNote] = []
    for sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        header_row = find_header_row(ws, "科目名称") or 29
        for row in range(header_row + 1, ws.max_row + 1):
            account = cell_text(ws.cell(row, 4).value)
            if not account or account in {"合计", "总计"}:
                continue
            if any(keyword in account for keyword in LEGAL_KEYWORDS):
                notes.append(
                    note(
                        "Medium",
                        file,
                        sheet,
                        f"D{row}",
                        "特殊费用性质提示",
                        f"发现可能涉及法律、诉讼、咨询或中介性质的费用：“{account}”。",
                        "判断是否需要执行 VD.01.3 法律费用复核，或在底稿中说明不适用依据。",
                        "SOP - 复核法律费用/特殊费用",
                    )
                )
    return notes


def review_legal_program_consistency(wb, file: Path, special_notes: list[ReviewNote]) -> list[ReviewNote]:
    if "汇总" not in wb.sheetnames or not special_notes:
        return []
    ws = wb["汇总"]
    notes: list[ReviewNote] = []
    for row in range(2, ws.max_row + 1):
        procedure = cell_text(ws.cell(row, 3).value) + " " + cell_text(ws.cell(row, 6).value)
        execution = cell_text(ws.cell(row, 7).value)
        reason = cell_text(ws.cell(row, 8).value)
        if "法律费用" in procedure and execution == "否":
            examples = "、".join(n.review_note.split("：“")[-1].rstrip("”。") for n in special_notes[:4])
            notes.append(
                note(
                    "High",
                    file,
                    "汇总",
                    f"G{row}:H{row}",
                    "程序拒绝与底稿内容可能不一致",
                    f"汇总页选择不执行法律费用复核，但 BKD 中存在特殊费用线索：{examples}。",
                    "重新筛选法律、诉讼、咨询、中介类费用；判断 VD.01.3 是否应执行。如仍不执行，补充更具体依据。",
                    "SOP - VD.01.3 复核法律费用",
                )
            )
            if "无" in reason and any(keyword in examples for keyword in LEGAL_KEYWORDS):
                notes.append(
                    note(
                        "Medium",
                        file,
                        "汇总",
                        f"H{row}",
                        "拒绝理由可能过于笼统",
                        f"不执行理由为“{reason}”，但底稿存在相关费用名称，理由可能不足。",
                        "将理由改为基于筛选结果、金额阈值、费用性质和替代程序的具体说明。",
                        "SOP - 拒绝 PSP 易错点",
                    )
                )
    return notes


def find_header_row(ws, header_text: str) -> int | None:
    for row in range(1, min(ws.max_row, 80) + 1):
        for col in range(1, min(ws.max_column, 35) + 1):
            if cell_text(ws.cell(row, col).value) == header_text:
                return row
    return None


def review_bkd_notes(wb, file: Path) -> list[ReviewNote]:
    notes: list[ReviewNote] = []
    for sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        header_row = find_header_row(ws, "科目名称")
        if not header_row:
            continue
        for row in range(header_row + 1, ws.max_row + 1):
            account = cell_text(ws.cell(row, 4).value)
            if not account or account in {"合计", "总计"}:
                continue
            note_index = cell_text(ws.cell(row, 15).value)
            fluctuation_flag = cell_text(ws.cell(row, 16).value)
            qualitative_flag = cell_text(ws.cell(row, 17).value)
            if (fluctuation_flag == "是" or qualitative_flag == "是") and not note_index:
                notes.append(
                    note(
                        "High",
                        file,
                        sheet,
                        f"O{row}",
                        "异常项目未添加 Notes 索引",
                        f"“{account}”被标记为需进一步调查，但 Notes 索引为空。",
                        "在 O 列补充 Note 索引，并在下方 Notes 中记录 ARP 分析、原因、证据索引和结论。",
                        "SOP - BKD 异常波动和 Notes",
                    )
                )
    return notes


def review_cutoff(wb, file: Path) -> list[ReviewNote]:
    sheet = "VC&VD.01.4 截止性测试"
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    notes: list[ReviewNote] = []
    if is_blank(ws["B10"].value) or is_blank(ws["C10"].value):
        notes.append(
            note(
                "High",
                file,
                sheet,
                "B10:C10",
                "截止性测试期间依据缺失",
                "截止性测试期间或确定依据未完整填写。",
                "补充前后测试天数和确定依据，说明与费用确认时点及关账流程的关系。",
                "SOP - 截止性测试",
            )
        )

    sample_start = find_header_row(ws, "凭证编号")
    if not sample_start:
        return notes

    for row in range(sample_start + 1, ws.max_row + 1):
        voucher = cell_text(ws.cell(row, 2).value)
        if voucher == "Notes":
            break
        if not voucher or voucher.startswith("本期") or voucher.startswith("*"):
            continue
        amount = cell_text(ws.cell(row, 4).value)
        if as_number(ws.cell(row, 4).value) is None:
            continue
        evidence_name = cell_text(ws.cell(row, 6).value)
        transaction_date = cell_text(ws.cell(row, 8).value)
        current_period = cell_text(ws.cell(row, 9).value)
        subsequent_period = cell_text(ws.cell(row, 10).value)
        ey_works = cell_text(ws.cell(row, 11).value)
        if evidence_name.upper() in {"", "N/A", "NA"}:
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    f"F{row}",
                    "截止性支持证据缺失",
                    f"凭证 {voucher}（金额 {amount}）未填写有效的支持性证据信息。",
                    "补充发票、报销单、合同、银行回单、明细表等支持性证据名称。",
                    "SOP - 截止性测试属性",
                )
            )
        if not transaction_date:
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    f"H{row}",
                    "交易发生日期缺失",
                    f"凭证 {voucher} 未填写交易发生日期，无法判断是否记录在适当期间。",
                    "根据支持性证据补充交易发生日期，并判断本期/期后期间归属。",
                    "SOP - 截止性测试属性",
                )
            )
        if current_period != "√" and subsequent_period != "√":
            notes.append(
                note(
                    "Medium",
                    file,
                    sheet,
                    f"I{row}:J{row}",
                    "期间归属标记缺失",
                    f"凭证 {voucher} 未标记本期或资产负债表日后期间。",
                    "在 I/J 列标记交易期间归属。",
                    "SOP - 截止性测试属性",
                )
            )
        if ey_works.upper() != "Y":
            notes.append(
                note(
                    "Medium",
                    file,
                    sheet,
                    f"K{row}",
                    "EY Works 结论缺失",
                    f"凭证 {voucher} 的 EY Works 列未标记 Y。",
                    "完成测试后在 K 列标记 Y；如存在例外，补充备注说明。",
                    "SOP - 截止性测试属性",
                )
            )
    return notes


def review_tod(wb, file: Path) -> list[ReviewNote]:
    notes: list[ReviewNote] = []
    visible_sheets = [s for s in wb.sheetnames if not s.startswith(INTERNAL_SHEET_PREFIXES)]
    total_pool_sheets = [s for s in visible_sheets if "总样本池" in s]
    reduced_pool_sheets = [s for s in visible_sheets if "剔除后样本池" in s]
    skywind_sheets = [s for s in visible_sheets if "Skywind" in s or "抽样工具输出" in s]
    test_sheets = [s for s in visible_sheets if ("测试" in s or "明细" in s or "TOD" in s or "样本" in s) and "样本池" not in s and "抽样" not in s]

    if not total_pool_sheets:
        notes.append(
            note(
                "High",
                file,
                "Workbook",
                "Sheet list",
                "TOD 总样本池缺失",
                "TOD 底稿中未识别到总样本池 sheet。",
                "补充或保留总样本池，以便复核样本总体完整性。",
                "SOP - TOD 样本总体",
            )
        )
    if not reduced_pool_sheets:
        notes.append(
            note(
                "Medium",
                file,
                "Workbook",
                "Sheet list",
                "TOD 剔除后样本池缺失",
                "TOD 底稿中未识别到剔除后样本池 sheet。",
                "补充剔除后样本池，并记录剔除逻辑和金额核对。",
                "SOP - TOD 剔除样本",
            )
        )
    if not skywind_sheets:
        notes.append(
            note(
                "Medium",
                file,
                "Workbook",
                "Sheet list",
                "抽样工具输出缺失",
                "TOD 底稿中未识别到 Skywind 或抽样工具输出 sheet。",
                "保留抽样工具输出，或在底稿中说明未使用抽样工具并记录抽样过程。",
                "SOP - TOD 抽样工具输出",
            )
        )
    if not test_sheets:
        notes.append(
            note(
                "High",
                file,
                "Workbook",
                "Sheet list",
                "TOD 测试明细或结论缺失",
                "TOD 底稿中未识别到样本测试明细或测试结论相关 sheet。",
                "补充样本测试明细、支持性证据检查结果和测试结论；如 sheet 命名不标准，请确保标题区域能清楚说明用途。",
                "SOP - TOD 样本测试",
            )
        )

    for sheet in total_pool_sheets + reduced_pool_sheets:
        ws = wb[sheet]
        if ws.max_row <= 1:
            notes.append(
                note(
                    "High",
                    file,
                    sheet,
                    "A1",
                    "样本池无明细数据",
                    f"{sheet} 未识别到明细行。",
                    "确认序时账筛选结果是否已粘贴或链接至该样本池。",
                    "SOP - TOD 样本池",
                )
            )
        elif ws.max_row < 3:
            notes.append(
                note(
                    "Medium",
                    file,
                    sheet,
                    "A1",
                    "TOD 样本池明细偏少",
                    f"{sheet} 仅识别到少量行，需确认样本池是否完整粘贴。",
                    "核对样本池是否包含完整筛选结果、金额字段和抽样范围。",
                    "SOP - TOD 样本池完整性",
                )
            )
    for sheet in test_sheets[:5]:
        ws = wb[sheet]
        text_hit = False
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 80), values_only=True):
            row_text = " ".join(cell_text(v) for v in row)
            if any(token in row_text for token in ("结论", "异常", "支持性", "EY Works", "测试结果")):
                text_hit = True
                break
        if not text_hit:
            notes.append(
                note(
                    "Medium",
                    file,
                    sheet,
                    "A1",
                    "TOD 测试结论线索不足",
                    f"{sheet} 前 80 行未识别到测试结论、异常说明或支持性证据字段。",
                    "确认 TOD 样本测试页是否已填写支持性证据、测试结果、异常处理和最终结论。",
                    "SOP - TOD 样本测试结论",
                )
            )
    return notes


# ---------------------------------------------------------------------------
#  NEW rules from Plan B
# ---------------------------------------------------------------------------

# 标准 VC&VD 底稿应包含的 sheet 集合
STANDARD_SHEETS_MAIN = {
    "汇总",
    "Uexp.00 Lead",
    "VC.00 销售费用BKD",
    "VD.00 管理费用BKD",
    "VC&VD.01.2 详细测试 TOD",
    "VD.01.3 复核法律费用",
    "VC&VD.01.4 截止性测试",
}
NON_EXPENSE_POPULATIONS = ("制造费用", "研发支出")


def review_bkd_currency_unit(wb, file: Path) -> list[ReviewNote]:
    """检查 BKD 表头货币/单位是否显示为 IFRS 而非 CNY。"""
    notes: list[ReviewNote] = []
    lead_currency = ""
    if "Uexp.00 Lead" in wb.sheetnames:
        lead_currency = cell_text(wb["Uexp.00 Lead"]["C8"].value).upper()

    for sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        currency_label = cell_text(ws.cell(28, 5).value)
        if "IFRS" in currency_label.upper():
            notes.append(
                note(
                    "High", file, sheet, "Row 28",
                    "BKD 货币/单位显示异常",
                    f"{sheet} 表头「货币/单位」显示为「IFRS」，IFRS 是会计准则而非货币。Lead 页 C8 记录为「{lead_currency}」。",
                    "将货币/单位更正为 CNY/RMB/人民币，与 Lead C8 保持一致。",
                    "SOP - Lead 基础信息 + BKD 易错点",
                )
            )
    return notes


def review_bkd_prior_year(wb, file: Path) -> list[ReviewNote]:
    """检测上期审定数列是否出现极小比例值（疑似公式错误）。"""
    notes: list[ReviewNote] = []
    for sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        header_row = find_header_row(ws, "科目名称")
        if not header_row:
            continue
        suspicious_rows: list[str] = []
        for row in range(header_row + 1, min(header_row + 30, ws.max_row + 1)):
            py_val = ws.cell(row, 10).value   # 上期审定数 (J列)
            cy_val = ws.cell(row, 7).value    # 本期审定数 (G列)
            # 上期值为极小比例(<0.1)但本期金额>100，大概率是公式 link 到了结构比列
            if (py_val is not None and isinstance(py_val, float) and abs(py_val) < 0.1
                    and cy_val is not None and isinstance(cy_val, (int, float)) and abs(cy_val) > 100):
                account = cell_text(ws.cell(row, 4).value)
                suspicious_rows.append(f"D{row} {account}")
        if len(suspicious_rows) >= 3:
            notes.append(
                note(
                    "Medium", file, sheet, "上期审定数列",
                    "上期审定数疑似公式错误",
                    f"{sheet} 中存在 {len(suspicious_rows)} 行上期审定数为极小比例值（<0.1），疑似公式 link 到了结构比列。行：{', '.join(suspicious_rows[:5])}。",
                    "检查上期审定数公式/链接，确保指向的是上期实际金额而非结构比。与上年底稿期末审定数核对一致。",
                    "SOP - BKD 易错点：若需插行时相关列公式未拉全",
                )
            )
    return notes


def _find_tod_detail_sheet(wb) -> str | None:
    """在 workbook 中查找 TOD 明细 sheet（包含预审/剩余期间等关键词）。"""
    for s in wb.sheetnames:
        if not s.startswith(INTERNAL_SHEET_PREFIXES):
            if "TOD" in s and ("预审" in s or "剩余期间" in s or "测试" in s):
                return s
    return None


def review_tod_population_scope(tod_wb, tod_path: Path) -> list[ReviewNote]:
    """检查 TOD 样本总体是否混入了制造费用/研发支出。"""
    notes: list[ReviewNote] = []
    if not tod_wb:
        return notes
    sheet = _find_tod_detail_sheet(tod_wb)
    if not sheet:
        return notes
    ws = tod_wb[sheet]
    for row in range(5, min(ws.max_row + 1, 25)):
        name = cell_text(ws.cell(row, 4).value)
        if any(pop in name for pop in NON_EXPENSE_POPULATIONS):
            amount = cell_text(ws.cell(row, 6).value)
            notes.append(
                note(
                    "High", tod_path, sheet, f"D{row}",
                    "样本总体包含非同类交易类别",
                    f"TOD 样本总体中包含「{name}」（金额约 {amount}），SOP 禁止将风险不同的制造费用/研发支出与销售/管理费用合并抽样。",
                    "将制造费用和研发支出从费用 TOD 样本总体中剔除，在对应科目底稿中单独执行 TOD。",
                    "SOP - TOD 易错点：将风险不同的制造费用与销售费用、管理费用合并抽样",
                )
            )
    return notes


def review_tod_key_items(tod_wb, tod_path: Path) -> list[ReviewNote]:
    """检查 TOD 关键项(KI)数量和金额是否全为 0。"""
    notes: list[ReviewNote] = []
    if not tod_wb:
        return notes
    sheet = _find_tod_detail_sheet(tod_wb)
    if not sheet:
        return notes
    ws = tod_wb[sheet]
    # 在关键项区域附近扫描数值
    ki_values: list[float] = []
    for row in range(30, min(ws.max_row + 1, 55)):
        for col in (6, 8):
            v = as_number(ws.cell(row, col).value)
            if v and v != 0:
                ki_values.append(v)
    if not ki_values:
        notes.append(
            note(
                "High", tod_path, sheet, "关键项区域",
                "TOD 关键项(KI)全部为零",
                "TOD 底稿中定量/定性关键项数量和金额均为 0。根据 SOP，性质特殊的费用（如法律/咨询/中介费用）应选为定性关键项单独检查。",
                "1) 确认测试阈值（取发生/计量/列报认定中最小 TT）；2) 将 >TT 的项选为定量 KI；3) 将咨询费、中介服务费等性质特殊费用选为定性 KI。",
                "SOP - TOD 关键项易错点：性质特殊的项未作为关键样本",
            )
        )
    return notes


def review_tod_negative_handling(tod_wb, tod_path: Path) -> list[ReviewNote]:
    """检查负值分析中「是否抽样」决策是否缺失。"""
    notes: list[ReviewNote] = []
    if not tod_wb:
        return notes
    sheet = _find_tod_detail_sheet(tod_wb)
    if not sheet:
        return notes
    ws = tod_wb[sheet]
    # 扫描负值分析区域，最多输出 2 条
    neg_note_count = 0
    for row in range(70, min(ws.max_row + 1, 105)):
        category = cell_text(ws.cell(row, 3).value)
        amount_val = cell_text(ws.cell(row, 4).value)
        decision = cell_text(ws.cell(row, 5).value)
        if not category or not amount_val:
            continue
        # 仅当分类名称包含"性质"、有金额、但无抽样决策时触发
        if "性质" in category and not decision and neg_note_count < 2:
            notes.append(
                note(
                    "High", tod_path, sheet, f"D{row}",
                    "负值处理决策缺失",
                    f"负值分析表「{category}」金额 {amount_val} 的「是否抽样」列为空，跳过决策步骤。",
                    "根据负值绝对值是否超过 TT、性质是否特殊，判断是否需抽样检查，并记录决策。",
                    "SOP - TOD 负值检查易错点：未对已剔除的金额重大或性质特殊的负值进行检查",
                )
            )
            neg_note_count += 1
    return notes


def review_missing_sheets(wb, file: Path) -> list[ReviewNote]:
    """对照标准模板检查主底稿是否缺失关键 sheet，并与汇总页执行状态联动。"""
    notes: list[ReviewNote] = []
    visible = {s for s in wb.sheetnames if not s.startswith(INTERNAL_SHEET_PREFIXES)}
    # 判断是否为主底稿（有 Lead 或 汇总）
    if "Uexp.00 Lead" not in wb.sheetnames and "汇总" not in wb.sheetnames:
        return notes

    # 读取汇总页执行状态，用于联动判断风险等级
    summary_executed: dict[str, str] = {}  # sheet_name -> 执行状态("是"/"否"/"")
    if "汇总" in wb.sheetnames:
        ws = wb["汇总"]
        for row in range(2, ws.max_row + 1):
            procedure = cell_text(ws.cell(row, 6).value)  # F列：程序页
            execution = cell_text(ws.cell(row, 7).value)   # G列：是否执行
            if procedure:
                summary_executed[procedure] = execution

    missing = STANDARD_SHEETS_MAIN - visible
    # sheet 名称到汇总页程序名的映射
    sheet_to_program = {
        "VC&VD.01.2 详细测试 TOD": "VC&VD.01.2 详细测试 TOD",
        "VD.01.3 复核法律费用": "VD.01.3 复核法律费用",
        "VC&VD.01.4 截止性测试": "VC&VD.01.4 截止性测试",
    }

    for sheet in sorted(missing):
        # 汇总页联动：如果缺失的sheet在汇总页标记为"是"执行，风险升级为 High
        program_name = sheet_to_program.get(sheet, "")
        exec_status = summary_executed.get(sheet, summary_executed.get(program_name, ""))
        is_marked_executed = (exec_status == "是")

        if sheet in ("汇总", "Uexp.00 Lead"):
            # 核心结构页缺失 -> High
            notes.append(
                note(
                    "High", file, sheet, "Workbook",
                    "缺少核心底稿页",
                    f"主底稿缺少核心结构页「{sheet}」，无法完成对应程序的 Review。",
                    f"补充「{sheet}」sheet；确认是否使用正确的标准底稿模板。",
                    "标准底稿模板结构",
                )
            )
        elif sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
            # BKD 页缺失 -> High（费用分析的基础）
            notes.append(
                note(
                    "High", file, sheet, "Workbook",
                    "缺少费用 BKD 页",
                    f"主底稿缺少「{sheet}」，无法执行费用波动分析程序。",
                    f"补充「{sheet}」sheet，或确认该费用类别是否在本审计范围内。",
                    "标准底稿模板 + SOP BKD 分析",
                )
            )
        elif sheet in ("VC&VD.01.2 详细测试 TOD",):
            risk = "High" if is_marked_executed else "Medium"
            suffix = f"且汇总页标记该程序为「是」执行" if is_marked_executed else ""
            notes.append(
                note(
                    risk, file, sheet, "Workbook",
                    "缺少标准底稿页",
                    f"主底稿缺少「{sheet}」{suffix}。标准底稿模板包含此页。",
                    f"补充「{sheet}」sheet，或确认程序是否在其他底稿（如 TOD 底稿）中执行并建立交叉索引。",
                    "标准底稿模板 + SOP 汇总页易错点",
                )
            )
        elif sheet in ("VD.01.3 复核法律费用",):
            risk = "High" if is_marked_executed else "Medium"
            suffix = f"且汇总页标记该程序为「是」执行" if is_marked_executed else ""
            notes.append(
                note(
                    risk, file, sheet, "Workbook",
                    "缺少标准底稿页",
                    f"主底稿缺少「{sheet}」{suffix}。即使程序不执行，SOP 也建议保留该页并记录原因。",
                    f"补充「{sheet}」sheet，如不执行则记录具体理由和替代程序依据。",
                    "标准底稿模板 + SOP 汇总页易错点",
                )
            )
        elif sheet in ("VC&VD.01.4 截止性测试",):
            risk = "High" if is_marked_executed else "Medium"
            suffix = f"且汇总页标记该程序为「是」执行" if is_marked_executed else ""
            notes.append(
                note(
                    risk, file, sheet, "Workbook",
                    "缺少标准底稿页",
                    f"主底稿缺少「{sheet}」{suffix}。截止性测试是费用审计的关键程序。",
                    f"补充「{sheet}」sheet 并完成截止性测试，或在汇总页说明不执行的具体理由。",
                    "标准底稿模板 + SOP 截止性测试",
                )
            )
    return notes


def review_missing_tod(main_path: Path) -> list[ReviewNote]:
    return [
        note(
            "High",
            main_path,
            "Workbook",
            "Uploaded files",
            "TOD 底稿未提供",
            "本次上传的一套底稿中未识别到 TOD 抽样底稿，无法复核主底稿中 TOD 程序的抽样总体、样本及测试结论。",
            "请补充同一套底稿中的 TOD 抽样底稿；如该项目不适用 TOD，应在汇总页和相关程序页说明不适用依据。",
            "上传文件识别 + SOP - TOD 程序",
        )
    ]


def make_summary(notes: list[ReviewNote], program_rows: int, checklist_rows: int) -> ReviewSummary:
    return ReviewSummary(
        total_notes=len(notes),
        high_count=sum(1 for n in notes if n.risk_level == "High"),
        medium_count=sum(1 for n in notes if n.risk_level == "Medium"),
        low_count=sum(1 for n in notes if n.risk_level == "Low"),
        sheets_impacted=len({(n.file, n.sheet) for n in notes}),
        rules_run=12,
        program_rows=program_rows,
        checklist_rows=checklist_rows,
    )


def review_files(
    files: dict[str, object] | None = None,
    *,
    llm_config: LLMConfig | None = None,
) -> tuple[ReviewSummary, list[ReviewNote]]:
    paths = dict(DEFAULT_FILES)
    if files:
        for key, value in files.items():
            if not value:
                continue
            if key == "workpapers":
                paths[key] = [Path(p) for p in value]  # type: ignore[arg-type]
            else:
                paths[key] = Path(value)  # type: ignore[arg-type]

    notes: list[ReviewNote] = []
    selection = select_workpapers(paths)
    main_path = selection.main_workpaper
    tod_path = selection.tod_workpaper

    main_wb = load_workbook(main_path)

    notes.extend(review_lead_basics(main_wb, main_path))
    notes.extend(review_summary_execution(main_wb, main_path))
    special_notes = find_special_expenses(main_wb, main_path)
    notes.extend(special_notes)
    notes.extend(review_legal_program_consistency(main_wb, main_path, special_notes))
    notes.extend(review_bkd_notes(main_wb, main_path))
    notes.extend(review_bkd_currency_unit(main_wb, main_path))
    notes.extend(review_bkd_prior_year(main_wb, main_path))
    notes.extend(review_cutoff(main_wb, main_path))
    if tod_path:
        tod_wb = load_workbook(tod_path)
        notes.extend(review_tod(tod_wb, tod_path))
        notes.extend(review_tod_population_scope(tod_wb, tod_path))
        notes.extend(review_tod_key_items(tod_wb, tod_path))
        notes.extend(review_tod_negative_handling(tod_wb, tod_path))
    else:
        notes.extend(review_missing_tod(main_path))
    notes.extend(review_missing_sheets(main_wb, main_path))

    program_rows = read_program_rows(Path(paths["program_doc"]))
    checklist_rows = read_review_checklist_rows(Path(paths["sop_workbook"]))
    return make_summary(notes, program_rows, checklist_rows), notes


def notes_to_rows(notes: Iterable[ReviewNote]) -> list[dict[str, str]]:
    return [asdict(n) for n in notes]


def notes_to_csv(notes: Iterable[ReviewNote]) -> bytes:
    rows = notes_to_rows(notes)
    if not rows:
        return b""
    import io

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue().encode("utf-8-sig")


if __name__ == "__main__":
    summary, review_notes = review_files()
    print(asdict(summary))
    for item in notes_to_rows(review_notes):
        print(item)
