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

from llm_assistant import LLMConfig, call_deepseek_json, confidence_text

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
    "trial_balance": DATA_DIR / "1253、202512科目余额表（更新） (1).xlsx",
    "general_ledger": DATA_DIR / "1253、202512序时账（更新） (1).xlsx",
    "program_doc": DATA_DIR / "审计程序要求.docx",
    "sop_workbook": SOP_DIR / "FY26_SOP U_exp SWP VC&VD.xlsx",
}

LEGAL_KEYWORDS = ("法律", "律师", "诉讼", "咨询", "中介")
INTERNAL_SHEET_PREFIXES = ("Skywind", "DS_INTERNAL")
GL_KEYWORDS = LEGAL_KEYWORDS + ("赔偿", "罚款", "处罚", "关联方", "调整", "暂估", "冲销", "补提")
MAX_LLM_CANDIDATES = 30


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


def collect_bkd_accounts(wb) -> set[str]:
    accounts: set[str] = set()
    for sheet in ("VC.00 销售费用BKD", "VD.00 管理费用BKD"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        header_row = find_header_row(ws, "科目名称") or 29
        for row in range(header_row + 1, ws.max_row + 1):
            account = cell_text(ws.cell(row, 4).value)
            if account and account not in {"合计", "总计"}:
                accounts.add(account)
    return accounts


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


def llm_notes_from_candidates(
    *,
    candidates: list[dict[str, object]],
    config: LLMConfig | None,
    task: str,
    default_source: str,
) -> tuple[list[ReviewNote], dict]:
    """返回 (notes, llm_status)。

    llm_status 格式: {"ok": bool, "status": str, "candidates_submitted": int}
    失败信息不再混入 Review Notes，由调用方在 UI 中单独展示。
    """
    base_status = {"ok": False, "status": "LLM not called", "candidates_submitted": len(candidates)}
    if not config or not config.is_available or not candidates:
        return [], base_status
    selected = candidates[:MAX_LLM_CANDIDATES]
    judgements, api_status = call_deepseek_json(config=config, task=task, candidates=selected)
    by_id = {str(item.get("candidate_id")): item for item in judgements}
    notes: list[ReviewNote] = []
    for candidate in selected:
        item = by_id.get(str(candidate.get("candidate_id")))
        if not item or not item.get("should_raise_note"):
            continue
        confidence = item.get("confidence", "")
        risk_level = cell_text(item.get("risk_level")) or "Medium"
        try:
            conf_num = float(confidence)
        except Exception:
            conf_num = 0.0
        if conf_num < 0.65:
            risk_level = "Low"
        elif conf_num < 0.8 and risk_level == "High":
            risk_level = "Medium"
        notes.append(
            note(
                risk_level if risk_level in {"High", "Medium", "Low"} else "Medium",
                Path(cell_text(candidate.get("file"))),
                cell_text(candidate.get("sheet")),
                cell_text(candidate.get("location")),
                cell_text(item.get("issue_type")) or cell_text(candidate.get("issue_type")),
                cell_text(item.get("reason")) or cell_text(candidate.get("reason")),
                cell_text(item.get("suggested_action")) or "请结合底稿和支持性资料进一步复核该候选事项。",
                default_source,
                judge_method="LLM Assist",
                confidence=confidence_text(confidence),
                evidence_summary=cell_text(candidate.get("evidence_summary")),
            )
        )
    llm_ok = api_status == "LLM ok"
    return notes, {
        "ok": llm_ok,
        "status": api_status,
        "candidates_submitted": len(selected),
        "notes_from_llm": len(notes),
    }


def _is_expense_parent_account(account: str) -> bool:
    """判断一个科目是否为费用类的父级汇总科目（非叶子科目）。

    TB 中父级科目如"销售费用"、"管理费用"仅作为汇总行，
    BKD 中通常只列示明细科目，精确匹配会导致误报。
    """
    clean = account.strip().replace(" ", "")
    # 如果科目名称正好是"销售费用"或"管理费用"（不含下级），则为父级
    if clean in {"销售费用", "管理费用"}:
        return True
    # 如果包含分隔符(_、-、/)则大概率是叶子科目
    if any(sep in clean for sep in ("_", "-", "/", "（")):
        return False
    # 其他情况：判断是否有明显的明细后缀
    return len(clean) <= 6  # "销售费用_xxx" > 6 个字


def review_trial_balance(tb_path: Path | None, main_wb, main_path: Path, llm_config: LLMConfig | None = None) -> list[ReviewNote]:
    if not tb_path or not tb_path.exists():
        return [
            note(
                "Low",
                main_path,
                "辅助资料",
                "科目余额表",
                "未上传科目余额表",
                "未上传科目余额表，无法辅助验证销售费用、管理费用 BKD 科目列示是否完整。",
                "如现场资料允许，请上传科目余额表，用于比对 BKD 是否覆盖全部销售费用、管理费用科目。",
                "辅助资料完整性",
            )
        ]

    bkd_accounts = collect_bkd_accounts(main_wb)
    wb = load_workbook(tb_path)
    ws = wb[wb.sheetnames[0]]
    notes: list[ReviewNote] = []
    candidates: list[dict[str, object]] = []
    rule_note_count = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        account = cell_text(row[1] if len(row) > 1 else None)
        if not account or ("销售费用" not in account and "管理费用" not in account):
            continue
        # 父级汇总科目不直接报 Medium，仅进入候选列表供 LLM 辅助判断
        is_parent = _is_expense_parent_account(account)
        if is_parent:
            if len(candidates) < MAX_LLM_CANDIDATES * 3:
                candidates.append(
                    {
                        "candidate_id": f"tb-{row_idx}",
                        "file": tb_path.name,
                        "sheet": ws.title,
                        "location": f"B{row_idx}",
                        "issue_type": "TB/BKD 科目列示差异",
                        "account_name": account,
                        "amount": row[8] if len(row) > 8 else None,
                        "reason": f"TB 中存在汇总科目 {account}，BKD 中未精确匹配同名科目（可能为父级汇总行）。",
                        "evidence_summary": f"TB 行 {row_idx}：科目={account}（父级汇总），本年累计借方={row[8] if len(row) > 8 else None}",
                    }
                )
            continue
        if account not in bkd_accounts:
            amount = row[8] if len(row) > 8 else None
            if rule_note_count < 10:
                notes.append(
                    note(
                        "Medium",
                        tb_path,
                        ws.title,
                        f"B{row_idx}",
                        "TB 科目可能未在 BKD 列示",
                        f"科目余额表存在费用科目“{account}”，但未在销售费用/管理费用 BKD 科目列示中匹配到同名科目。",
                        "核对该科目是否应纳入 BKD；如属于重分类、合并列示或非审计范围，请在底稿中说明对应关系。",
                        "科目余额表 + BKD 科目列示",
                        evidence_summary=f"TB 行 {row_idx}：科目={account}，本年累计借方={amount}",
                    )
                )
                rule_note_count += 1
            if len(candidates) < MAX_LLM_CANDIDATES * 3:
                candidates.append(
                    {
                        "candidate_id": f"tb-{row_idx}",
                        "file": tb_path.name,
                        "sheet": ws.title,
                        "location": f"B{row_idx}",
                        "issue_type": "TB/BKD 科目列示差异",
                        "account_name": account,
                        "amount": amount,
                        "reason": f"TB 中存在 {account}，BKD 中未精确匹配同名科目。",
                        "evidence_summary": f"TB 行 {row_idx}：科目={account}，本年累计借方={amount}",
                    }
                )
    llm_notes, _ = llm_notes_from_candidates(
        candidates=candidates,
        config=llm_config,
        task="判断 TB 与销售费用/管理费用 BKD 的科目列示差异是否需要形成审计 Review Note。",
        default_source="DeepSeek + 科目余额表/BKD 候选证据",
    )
    notes.extend(llm_notes)
    return notes


def review_general_ledger(gl_path: Path | None, main_wb, main_path: Path, llm_config: LLMConfig | None = None) -> list[ReviewNote]:
    if not gl_path or not gl_path.exists():
        return [
            note(
                "Low",
                main_path,
                "辅助资料",
                "序时账",
                "未上传序时账",
                "未上传序时账，无法辅助验证是否存在特殊费用、大额异常费用或期末集中入账未被 BKD 分析覆盖。",
                "如现场资料允许，请上传序时账，用于补充识别法律、诉讼、咨询、中介、赔偿、罚款等费用线索。",
                "辅助资料完整性",
            )
        ]

    wb = load_workbook(gl_path)
    ws = wb[wb.sheetnames[0]]
    notes: list[ReviewNote] = []
    candidates: list[dict[str, object]] = []
    rule_note_count = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        summary = cell_text(row[5] if len(row) > 5 else None)
        account = cell_text(row[7] if len(row) > 7 else None)
        full_account = cell_text(row[8] if len(row) > 8 else None)
        amount = row[11] if len(row) > 11 else row[10] if len(row) > 10 else None
        period = row[2] if len(row) > 2 else None
        text = f"{summary} {account} {full_account}"
        if not any(keyword in text for keyword in GL_KEYWORDS):
            continue
        # 只对销售费用/管理费用类科目的分录生成 Review Note
        is_expense_account = ("销售费用" in (full_account or "") or "管理费用" in (full_account or "")
                              or "销售费用" in (account or "") or "管理费用" in (account or ""))
        evidence = f"序时账行 {row_idx}：期间={period}，摘要={summary}，科目={full_account or account}，金额={amount}"
        if is_expense_account and rule_note_count < 10:
            notes.append(
                note(
                    "Medium",
                    gl_path,
                    ws.title,
                    f"F{row_idx}:I{row_idx}",
                    "序时账特殊费用线索",
                    f"序时账发现可能需要进一步关注的费用线索：摘要“{summary}”，科目“{full_account or account}”。",
                    "核对该交易是否已在 BKD 或相关专项程序中分析；如涉及法律、诉讼、咨询、中介等性质，请补充判断依据和支持性证据索引。",
                    "序时账辅助扫描 + SOP 特殊费用",
                    evidence_summary=evidence,
                )
            )
            rule_note_count += 1
        if len(candidates) < MAX_LLM_CANDIDATES * 3:
            candidates.append(
                {
                    "candidate_id": f"gl-{row_idx}",
                    "file": gl_path.name,
                    "sheet": ws.title,
                    "location": f"F{row_idx}:I{row_idx}",
                    "issue_type": "序时账特殊费用线索",
                    "period": period,
                    "summary": summary,
                    "account_name": full_account or account,
                    "amount": amount,
                    "reason": "序时账摘要或科目命中特殊费用/异常关键词。",
                    "evidence_summary": evidence,
                }
            )
    llm_notes, _ = llm_notes_from_candidates(
        candidates=candidates,
        config=llm_config,
        task="判断序时账候选交易是否应触发费用底稿 Review Note，重点关注法律、诉讼、咨询、中介、赔偿、罚款、关联方、调整、暂估、冲销、补提等。",
        default_source="DeepSeek + 序时账候选证据",
    )
    notes.extend(llm_notes)
    return notes


def make_summary(notes: list[ReviewNote], program_rows: int, checklist_rows: int) -> ReviewSummary:
    return ReviewSummary(
        total_notes=len(notes),
        high_count=sum(1 for n in notes if n.risk_level == "High"),
        medium_count=sum(1 for n in notes if n.risk_level == "Medium"),
        low_count=sum(1 for n in notes if n.risk_level == "Low"),
        sheets_impacted=len({(n.file, n.sheet) for n in notes}),
        rules_run=8,
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
            if value is None and key in {"trial_balance", "general_ledger"}:
                paths[key] = None
                continue
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
    notes.extend(review_cutoff(main_wb, main_path))
    if tod_path:
        tod_wb = load_workbook(tod_path)
        notes.extend(review_tod(tod_wb, tod_path))
    else:
        notes.extend(review_missing_tod(main_path))
    notes.extend(review_trial_balance(paths.get("trial_balance"), main_wb, main_path, llm_config))  # type: ignore[arg-type]
    notes.extend(review_general_ledger(paths.get("general_ledger"), main_wb, main_path, llm_config))  # type: ignore[arg-type]

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
