from __future__ import annotations

import cgi
import io
import json
import os
import re
import sys
import time
import uuid
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
import xml.etree.ElementTree as ET


APP_DIR = Path(__file__).resolve().parent
ROOT_DIR = APP_DIR.parent
STATIC_DIR = APP_DIR / "static"
OUTPUT_DIR = APP_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)
DOWNLOADS: dict[str, bytes] = {}
AI_CONFIG = {
    "api_key": os.environ.get("OPENAI_API_KEY", "").strip(),
    "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini").strip(),
    "endpoint": os.environ.get("OPENAI_RESPONSES_URL", "https://api.openai.com/v1/responses").strip(),
}

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

ET.register_namespace("", NS["a"])
ET.register_namespace("r", NS["r"])
ET.register_namespace("rel", NS["rel"])


@dataclass
class Cell:
    value: Any = ""
    formula: str = ""


def text_of(node: ET.Element) -> str:
    return "".join(t.text or "" for t in node.iter() if t.tag.endswith("}t"))


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace(",", "").strip()
    if cleaned in {"-", "NA", "N/A"}:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def col_to_num(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref)
    if not match:
        return 0
    result = 0
    for ch in match.group(1):
        result = result * 26 + ord(ch) - 64
    return result


class XLSXBook:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self.data = data
        self.sheets: list[str] = []
        self.cells: dict[str, dict[str, Cell]] = {}
        self._parse()

    def _parse(self) -> None:
        with zipfile.ZipFile(io.BytesIO(self.data)) as zf:
            shared = self._shared_strings(zf)
            rels = self._workbook_rels(zf)
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
            for sheet in workbook.findall(".//a:sheet", NS):
                sheet_name = sheet.attrib["name"]
                rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                path = rels.get(rid)
                if not path or path not in zf.namelist():
                    continue
                self.sheets.append(sheet_name)
                self.cells[sheet_name] = self._sheet_cells(zf, path, shared)

    def _shared_strings(self, zf: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in zf.namelist():
            return []
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        return [text_of(si) for si in root.findall("a:si", NS)]

    def _workbook_rels(self, zf: zipfile.ZipFile) -> dict[str, str]:
        root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels: dict[str, str] = {}
        for rel in root.findall("rel:Relationship", NS):
            target = rel.attrib["Target"]
            if target.startswith("/"):
                path = target.lstrip("/")
            elif target.startswith("xl/"):
                path = target
            else:
                path = f"xl/{target}"
            rels[rel.attrib["Id"]] = path
        return rels

    def _sheet_cells(self, zf: zipfile.ZipFile, path: str, shared: list[str]) -> dict[str, Cell]:
        sheet = ET.fromstring(zf.read(path))
        cells: dict[str, Cell] = {}
        for node in sheet.findall(".//a:c", NS):
            ref = node.attrib.get("r", "")
            if not ref:
                continue
            formula_node = node.find("a:f", NS)
            value_node = node.find("a:v", NS)
            cell_type = node.attrib.get("t")
            formula = formula_node.text or "" if formula_node is not None else ""
            value: Any = ""
            if value_node is not None and value_node.text is not None:
                raw = value_node.text
                if cell_type == "s":
                    idx = int(raw)
                    value = shared[idx] if idx < len(shared) else ""
                else:
                    value = raw
            elif cell_type == "inlineStr":
                value = text_of(node)
            if formula_node is not None and value == "":
                value = f"={formula}"
            cells[ref] = Cell(value=value, formula=formula)
        return cells

    def find_sheet(self, *keywords: str) -> str | None:
        for sheet in self.sheets:
            if all(keyword in sheet for keyword in keywords):
                return sheet
        return None

    def cell(self, sheet: str | None, ref: str) -> Any:
        if not sheet:
            return ""
        return self.cells.get(sheet, {}).get(ref, Cell()).value

    def formula(self, sheet: str | None, ref: str) -> str:
        if not sheet:
            return ""
        return self.cells.get(sheet, {}).get(ref, Cell()).formula

    def rows(self, sheet: str | None) -> dict[int, dict[int, Any]]:
        output: dict[int, dict[int, Any]] = {}
        if not sheet:
            return output
        for ref, cell in self.cells.get(sheet, {}).items():
            match = re.match(r"([A-Z]+)(\d+)", ref)
            if not match:
                continue
            row = int(match.group(2))
            output.setdefault(row, {})[col_to_num(ref)] = cell.value
        return output


def xlsx_sheet_names_and_paths(data: bytes) -> list[tuple[str, str]]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        rels = {}
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        for rel in rel_root.findall("rel:Relationship", NS):
            target = rel.attrib["Target"]
            if target.startswith("/"):
                path = target.lstrip("/")
            elif target.startswith("xl/"):
                path = target
            else:
                path = f"xl/{target}"
            rels[rel.attrib["Id"]] = path
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        output = []
        for sheet in workbook.findall(".//a:sheet", NS):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            output.append((sheet.attrib["name"], rels[rid]))
        return output


def xlsx_shared_strings_from_zip(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [text_of(si) for si in root.findall("a:si", NS)]


def stream_first_sheet_rows(data: bytes):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared = xlsx_shared_strings_from_zip(zf)
        sheets = xlsx_sheet_names_and_paths(data)
        if not sheets:
            return
        sheet_name, sheet_path = sheets[0]
        with zf.open(sheet_path) as fp:
            for _event, elem in ET.iterparse(fp, events=("end",)):
                if not elem.tag.endswith("}row"):
                    continue
                row_idx = int(elem.attrib.get("r", "0") or 0)
                row: dict[int, Any] = {}
                for cell in elem:
                    if not cell.tag.endswith("}c"):
                        continue
                    ref = cell.attrib.get("r", "")
                    if not ref:
                        continue
                    value = ""
                    cell_type = cell.attrib.get("t")
                    value_node = None
                    for child in cell:
                        if child.tag.endswith("}v"):
                            value_node = child
                            break
                    if value_node is not None and value_node.text is not None:
                        raw = value_node.text
                        if cell_type == "s":
                            idx = int(raw)
                            value = shared[idx] if idx < len(shared) else ""
                        else:
                            value = raw
                    elif cell_type == "inlineStr":
                        value = text_of(cell)
                    row[col_to_num(ref)] = value
                yield sheet_name, row_idx, row
                elem.clear()


def parse_docx(name: str, data: bytes) -> dict[str, Any]:
    paragraphs: list[str] = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            document = ET.fromstring(zf.read("word/document.xml"))
            for p in document.findall(".//w:p", NS):
                text = text_of(p).strip()
                if text:
                    paragraphs.append(text)
    except Exception:
        pass
    return {"name": name, "kind": "program", "paragraphs": len(paragraphs), "preview": paragraphs[:8]}


def finding(
    findings: list[dict[str, Any]],
    severity: str,
    category: str,
    title: str,
    file_name: str,
    sheet: str,
    cell: str,
    evidence: str,
    review_note: str,
    recommendation: str,
    risk: str,
) -> None:
    findings.append(
        {
            "id": f"RN-{len(findings) + 1:03d}",
            "severity": severity,
            "category": category,
            "title": title,
            "location": {"file": file_name, "sheet": sheet, "cell": cell},
            "evidence": evidence,
            "reviewNote": review_note,
            "recommendation": recommendation,
            "risk": risk,
        }
    )


def identify_books(books: list[XLSXBook]) -> tuple[XLSXBook | None, XLSXBook | None]:
    main_book = None
    tod_book = None
    for book in books:
        if any("Uexp.00 Lead" in s for s in book.sheets):
            main_book = book
        if any("TOD" in s and ("预审" in s or "年审" in s) for s in book.sheets):
            if not any("Uexp.00 Lead" in s for s in book.sheets):
                tod_book = book
    return main_book, tod_book


def analyze_main_book(book: XLSXBook, tod_book: XLSXBook | None, findings: list[dict[str, Any]]) -> dict[str, Any]:
    lead = book.find_sheet("Uexp.00 Lead")
    summary = book.find_sheet("汇总") or (book.sheets[0] if book.sheets else None)
    sales = book.find_sheet("销售费用", "BKD")
    admin = book.find_sheet("管理费用", "BKD")
    cutoff = book.find_sheet("截止性测试")
    context = {
        "customer": as_text(book.cell(lead, "C2")),
        "period": as_text(book.cell(lead, "C3")),
        "te": book.cell(lead, "C5"),
        "sad": book.cell(lead, "C6"),
    }

    c7 = as_text(book.cell(lead, "C7"))
    c8 = as_text(book.cell(lead, "C8"))
    currency_values = {"CNY", "RMB", "USD", "EUR", "HKD", "JPY"}
    gaap_markers = ("IFRS", "GAAP", "企业会计准则", "会计准则")
    if c7.upper() in currency_values and any(marker in c8 for marker in gaap_markers):
        finding(
            findings,
            "High",
            "基础信息",
            "适用会计准则与记账本位币疑似填反",
            book.name,
            lead or "",
            "C7:C8",
            f"C7={c7}; C8={c8}",
            "Lead页中适用会计准则填列为币种，记账本位币填列为会计准则，字段含义与填列内容不一致。",
            "请将适用会计准则修改为 IFRS/企业会计准则等适用准则，并将记账本位币修改为 CNY 等币种。",
            "基础信息错误可能影响底稿编制基础和后续引用判断。",
        )

    analysis_date = as_text(book.cell(lead, "C4"))
    period_end = as_text(book.cell(lead, "C3"))
    if analysis_date and period_end and analysis_date == period_end:
        finding(
            findings,
            "Low",
            "描述规范",
            "分析日期与期末日期相同，建议确认是否已更新",
            book.name,
            lead or "",
            "C4",
            f"期末={period_end}; 分析日期={analysis_date}",
            "Lead页分析日期与资产负债表日一致，需确认是否为实际底稿编制日期。",
            "如该日期并非实际分析日期，请更新为底稿编制或分析完成日期。",
            "日期未更新通常不直接导致错报，但会影响底稿时点记录的清晰度。",
        )

    summary_rows = book.rows(summary)
    sap_execute = as_text(summary_rows.get(8, {}).get(7))
    sap_xref = as_text(summary_rows.get(8, {}).get(6))
    has_sap_sheet = any("264" in s or "SAP" in s or "01.1" in s for s in book.sheets)
    if sap_execute == "是" and ("如执行" in sap_xref or not has_sap_sheet):
        finding(
            findings,
            "High",
            "程序完整性",
            "SAP程序选择执行但缺少明确底稿索引",
            book.name,
            summary or "",
            "G8/F8",
            f"G8={sap_execute}; F8={sap_xref}",
            "汇总页显示VC&VD.01.1实质性分析程序已执行，但程序页位置仍为提示性文字，未见清晰的264GL或SAP底稿索引。",
            "请补充SAP底稿索引；如未执行SAP，应将执行状态调整为否并说明原因，或直接执行TOD程序。",
            "程序执行状态与支持底稿不一致，可能导致审计程序执行证据不足。",
        )

    tod_execute = as_text(summary_rows.get(10, {}).get(7))
    if tod_execute == "是" and tod_book is None:
        finding(
            findings,
            "High",
            "程序完整性",
            "TOD程序已选择执行但未上传TOD底稿",
            book.name,
            summary or "",
            "G10",
            "汇总页TOD程序状态为“是”，但本次分析未识别到TOD底稿文件。",
            "请上传TOD SWP，或在汇总页中补充可追踪的TOD底稿索引。",
            "缺少详细测试底稿会影响费用发生、计量及列报认定的审计证据充分性。",
        )

    analyze_bkd_sheet(book, sales, "销售费用", findings)
    analyze_bkd_sheet(book, admin, "管理费用", findings)
    analyze_cutoff(book, cutoff, findings)
    return context


def analyze_bkd_sheet(book: XLSXBook, sheet: str | None, account_type: str, findings: list[dict[str, Any]]) -> None:
    if not sheet:
        return
    rows = book.rows(sheet)
    expectation = as_text(rows.get(18, {}).get(2))
    if not expectation:
        finding(
            findings,
            "Medium",
            "缺失内容",
            f"{account_type}BKD未填写账户波动预期",
            book.name,
            sheet,
            "B18",
            "波动预期为空",
            f"{account_type}BKD未记录当期对账户波动的预期及理由。",
            "请根据业务变化、预算、上期情况或其他审计证据补充预期及依据。",
            "未设定预期会削弱后续波动分析结论的可复核性。",
        )

    material_promotion_rows = []
    for row, cols in rows.items():
        name = as_text(cols.get(4))
        current_amount = to_number(cols.get(5)) or 0
        fluct_flag = as_text(cols.get(16))
        qualitative_flag = as_text(cols.get(17))
        note = as_text(cols.get(15))
        if row > 25 and name and name != "合计" and (fluct_flag == "是" or qualitative_flag == "是") and not note:
            finding(
                findings,
                "Medium",
                "缺失内容",
                f"{name}触发进一步调查但未填写Notes索引",
                book.name,
                sheet,
                f"O{row}",
                f"P列={fluct_flag}; Q列={qualitative_flag}; Notes为空",
                "BKD中该科目触发进一步调查条件，但未填写对应分析说明或支持性底稿索引。",
                "请补充Notes编号，并在波动说明区域记录调查程序、证据来源和结论。",
                "触发调查但未记录处理过程，可能导致异常波动未被充分解释。",
            )
        if "推广费" in name and abs(current_amount) > 50_000_000:
            material_promotion_rows.append((row, name, current_amount))

    if material_promotion_rows:
        first_row, first_name, first_amount = material_promotion_rows[0]
        finding(
            findings,
            "Medium",
            "审计风险",
            f"{account_type}中推广费金额重大，建议保留风险关注",
            book.name,
            sheet,
            f"D{first_row}:E{first_row}",
            f"{first_name} 本期金额约 {first_amount:,.2f}",
            "推广费金额重大且属于监管和舞弊风险关注度较高的费用性质，需确认相关分析、支持性文件和必要的实质性程序已充分执行。",
            "请保留推广费测算或检查底稿索引，并确认合同、发票、服务成果或其他支持性证据与账面记录一致。",
            "推广费可能涉及费用真实性、截止、商业合理性及潜在舞弊风险。",
        )

    for row, cols in rows.items():
        if as_text(cols.get(6)) == "Diff":
            for col_num, col_letter in [(7, "G"), (10, "J"), (11, "K")]:
                diff = to_number(cols.get(col_num))
                if diff is not None and abs(diff) > 0.01:
                    finding(
                        findings,
                        "High",
                        "勾稽核对",
                        f"{account_type}BKD与外部表勾稽存在差异",
                        book.name,
                        sheet,
                        f"{col_letter}{row}",
                        f"Diff={diff:,.2f}",
                        "BKD勾稽差异不为零，需确认本期未审数、审定数或上期审定数是否已与A3/试算表核对一致。",
                        "请追查差异来源，并更新取数、公式或引用底稿。",
                        "金额勾稽差异可能影响BKD金额完整性和准确性。",
                    )


def analyze_cutoff(book: XLSXBook, sheet: str | None, findings: list[dict[str, Any]]) -> None:
    if not sheet:
        return
    strategy = as_text(book.cell(sheet, "B13"))
    if "日前15天" in strategy and "日后15天" not in strategy and "后15天" not in strategy:
        finding(
            findings,
            "Medium",
            "描述规范",
            "截止性测试策略描述未覆盖资产负债表日后期间",
            book.name,
            sheet,
            "B13",
            strategy[:180],
            "截止性测试策略描述仅提及资产负债表日前期间，但样本表同时包含日后记录的交易，描述范围与实际测试安排不完全一致。",
            "请将策略描述更新为资产负债表日前后各15天，或分别说明日前、日后样本选取标准。",
            "策略描述不完整可能影响测试范围和样本选取依据的可复核性。",
        )


def analyze_tod_book(book: XLSXBook, main_context: dict[str, Any], findings: list[dict[str, Any]]) -> None:
    tod_sheet = book.find_sheet("TOD", "预审") or book.find_sheet("TOD", "年审")
    if not tod_sheet:
        return
    strategy = as_text(book.cell(tod_sheet, "D49"))
    header_text = " ".join(as_text(book.cell(tod_sheet, ref)) for ref in ["F10", "G10"])
    if strategy:
        header_ranges = set(re.findall(r"\d+-\d+月", header_text))
        strategy_ranges = set(re.findall(r"\d+-\d+月", strategy))
        if header_ranges and strategy_ranges and header_ranges != strategy_ranges:
            finding(
                findings,
                "Medium",
                "逻辑一致性",
                "TOD预审和剩余期间描述与表头期间不一致",
                book.name,
                tod_sheet,
                "D49",
                f"表头期间={', '.join(sorted(header_ranges))}; 策略描述={', '.join(sorted(strategy_ranges))}",
                "TOD抽样策略中的期间描述与表头预审/剩余期间划分不一致。",
                "请统一预审期间和剩余期间的文字描述，确保与实际数据列一致。",
                "期间描述不一致可能导致样本分配和测试覆盖期间被误解。",
            )
        declared = re.search(r"(\d+)[^\d]{0,12}代表性样本", strategy)
        if declared:
            declared_count = int(declared.group(1))
            actual_count = 0
            for row, cols in book.rows(tod_sheet).items():
                if 55 <= row <= 150 and "代表性样本" in as_text(cols.get(4)):
                    actual_count += 1
            if actual_count and declared_count != actual_count:
                finding(
                    findings,
                    "High",
                    "样本完整性",
                    "TOD策略描述的代表性样本数量与测试表不一致",
                    book.name,
                    tod_sheet,
                    "D49",
                    f"策略描述={declared_count}个; 测试表识别={actual_count}个",
                    "TOD抽样策略记录的代表性样本数量与下方实际测试记录数量不一致。",
                    "请补充缺失样本测试记录，或修正抽样策略中的样本数量描述。",
                    "样本数量不一致会影响详细测试执行完整性和抽样结论。",
                )

    main_customer = as_text(main_context.get("customer"))
    if main_customer:
        for sheet in book.sheets:
            customer_label = as_text(book.cell(sheet, "A2"))
            customer_name = as_text(book.cell(sheet, "D2"))
            if customer_label == "客户名称" and customer_name and customer_name != main_customer:
                finding(
                    findings,
                    "High",
                    "引用一致性",
                    "TOD抽样底稿客户名称与主底稿不一致",
                    book.name,
                    sheet,
                    "D2",
                    f"主底稿客户={main_customer}; 当前表客户={customer_name}",
                    "TOD抽样输出中的客户名称与主底稿Lead页客户名称不一致。",
                    "请确认是否误用了其他项目或其他客户的抽样输出，并替换为本项目对应底稿。",
                    "客户名称不一致属于高风险引用错误，可能导致底稿证据不适用于当前项目。",
                )

    for row, cols in book.rows(tod_sheet).items():
        if 60 <= row <= 150 and as_text(cols.get(4)):
            for col_num, attr_name in [(16, "A"), (17, "B"), (18, "C")]:
                result = as_text(cols.get(col_num))
                if result.upper() == "N":
                    finding(
                        findings,
                        "High",
                        "测试异常",
                        f"TOD样本属性{attr_name}测试结果为N",
                        book.name,
                        tod_sheet,
                        f"{chr(64 + col_num)}{row}",
                        f"样本ID={as_text(cols.get(3))}; 属性{attr_name}=N",
                        "TOD样本属性测试存在未通过项目，但未在当前规则中识别到充分异常解释。",
                        "请在异常解释区域记录差异原因、追加程序及最终结论。",
                        "属性测试未通过可能影响费用发生、期间或分类认定。",
                    )


def analyze_tb_support(tb_name: str, tb_data: bytes, main_book: XLSXBook | None, findings: list[dict[str, Any]]) -> None:
    if not main_book:
        return
    tb_book = XLSXBook(tb_name, tb_data)
    sheet = tb_book.sheets[0] if tb_book.sheets else None
    if not sheet:
        return
    rows = tb_book.rows(sheet)
    totals = {}
    for row, cols in rows.items():
        code = as_text(cols.get(1))
        if code in {"6601", "6602"}:
            totals[code] = to_number(cols.get(15)) or to_number(cols.get(14)) or 0

    checks = [
        ("6601", main_book.find_sheet("销售费用", "BKD"), "销售费用", "E62"),
        ("6602", main_book.find_sheet("管理费用", "BKD"), "管理费用", "E88"),
    ]
    for code, bkd_sheet, label, cell in checks:
        if not bkd_sheet or code not in totals:
            continue
        bkd_total = to_number(main_book.cell(bkd_sheet, cell)) or 0
        diff = bkd_total - totals[code]
        if abs(diff) > 1:
            finding(
                findings,
                "High",
                "TB核对",
                f"{label}BKD合计与科目余额表不一致",
                main_book.name,
                bkd_sheet,
                cell,
                f"BKD={bkd_total:,.2f}; TB {code}={totals[code]:,.2f}; 差异={diff:,.2f}",
                f"{label}BKD合计与科目余额表对应科目金额存在差异。",
                "请检查BKD取数范围、科目映射、账表调整和公式是否正确，并更新差异说明。",
                "BKD与TB不一致会影响费用金额完整性和准确性。",
            )

    # Checklist style completeness check: if totals agree, still leave no finding; dashboard remains issue-driven.


def analyze_je_support(
    je_name: str,
    je_data: bytes,
    main_context: dict[str, Any],
    findings: list[dict[str, Any]],
    tod_book: XLSXBook | None,
    main_book: XLSXBook | None,
) -> None:
    sheet_name = ""
    tod_sheet = tod_book.find_sheet("TOD", "预审") or tod_book.find_sheet("TOD", "年审") if tod_book else None
    cutoff_sheet = main_book.find_sheet("截止性测试") if main_book else None
    tod_file = tod_book.name if tod_book and tod_sheet else je_name
    cutoff_file = main_book.name if main_book and cutoff_sheet else tod_file
    tod_location_sheet = tod_sheet or sheet_name
    cutoff_location_sheet = cutoff_sheet or tod_location_sheet
    sad = to_number(main_context.get("sad")) or 379000
    te = to_number(main_context.get("te")) or 3123450
    negative_large: list[tuple[int, str, str, str, float, str]] = []
    large_entries: list[tuple[int, str, str, str, float, str]] = []
    no_attachment: list[tuple[int, str, str, str, float, str]] = []
    cutoff_entries: list[tuple[int, str, str, str, float, str]] = []
    total_expense_lines = 0

    for sheet, row_idx, row in stream_first_sheet_rows(je_data):
        sheet_name = sheet_name or sheet
        if row_idx == 1:
            continue
        date = as_text(row.get(1))
        summary = as_text(row.get(6))
        code = as_text(row.get(7))
        name = as_text(row.get(9)) or as_text(row.get(8))
        if not (code.startswith("6601") or code.startswith("6602")):
            continue
        debit = to_number(row.get(12)) or 0
        credit = to_number(row.get(13)) or 0
        amount = debit - credit
        total_expense_lines += 1
        attachment = as_text(row.get(14))
        if amount < 0 and abs(amount) >= sad and len(negative_large) < 8:
            negative_large.append((row_idx, date, code, name, amount, summary))
        if abs(amount) >= te and len(large_entries) < 8:
            large_entries.append((row_idx, date, code, name, amount, summary))
        if not attachment and abs(amount) >= sad and len(no_attachment) < 8:
            no_attachment.append((row_idx, date, code, name, amount, summary))
        if (date.startswith("2025/12/") or date.startswith("2026/1/")) and abs(amount) >= sad and len(cutoff_entries) < 8:
            cutoff_entries.append((row_idx, date, code, name, amount, summary))

    if negative_large:
        row, date, code, name, amount, summary = negative_large[0]
        finding(
            findings,
            "Medium",
            "JE风险扫描",
            "序时账存在大额负数费用分录",
            tod_file,
            tod_location_sheet,
            "C69" if tod_sheet else f"L{row}:M{row}",
            f"识别到{len(negative_large)}条金额大于SAD的负数费用样本；示例：{date} {code} {name} {amount:,.2f} {summary[:80]}",
            "JE中存在大额负数费用分录，需确认是否为合理冲销、重分类或前期计提转回，并核对支持性证据。",
            "建议抽查大额负数费用分录，检查冲销依据、原始凭证、期间归属和金额计算。",
            "大额负数费用可能影响费用完整性、截止性及分类准确性。",
        )
    if large_entries:
        row, date, code, name, amount, summary = large_entries[0]
        finding(
            findings,
            "Medium",
            "JE风险扫描",
            "序时账存在超过TE的大额费用分录",
            tod_file,
            tod_location_sheet,
            "D35" if tod_sheet else f"L{row}:M{row}",
            f"识别到{len(large_entries)}条单笔绝对金额超过TE的费用分录；示例：{date} {code} {name} {amount:,.2f} {summary[:80]}",
            "JE中存在超过TE的大额费用分录，建议确认其是否已纳入关键项或TOD样本考虑。",
            "请与TOD抽样策略和关键项识别记录核对，必要时补充测试或说明未选取原因。",
            "大额费用未被纳入关键项考虑，可能导致详细测试覆盖不足。",
        )
    if no_attachment:
        row, date, code, name, amount, summary = no_attachment[0]
        finding(
            findings,
            "Low",
            "JE风险扫描",
            "部分大额费用分录附件数字段为空",
            tod_file,
            tod_location_sheet,
            "C69" if tod_sheet else f"N{row}",
            f"识别到{len(no_attachment)}条大于SAD且附件数字段为空的费用分录；示例：{date} {code} {name} {amount:,.2f} {summary[:80]}",
            "JE中部分大额费用分录附件数字段为空，需确认是否为分录拆行导致，或确实缺少支持性附件。",
            "请抽查相关凭证附件，确认合同、发票、报销单、银行回单等支持性文件是否完整。",
            "附件信息缺失会影响凭证检查和支持性证据追踪效率。",
        )
    if cutoff_entries:
        row, date, code, name, amount, summary = cutoff_entries[0]
        finding(
            findings,
            "Medium",
            "截止性风险",
            "期末附近存在大额费用分录，需与截止性测试样本核对",
            cutoff_file,
            cutoff_location_sheet,
            "B12" if cutoff_sheet else f"A{row}:M{row}",
            f"识别到{len(cutoff_entries)}条12月/1月且大于SAD的费用分录；示例：{date} {code} {name} {amount:,.2f} {summary[:80]}",
            "JE中存在期末附近的大额费用分录，需确认其是否已被截止性测试覆盖或有合理未选取原因。",
            "请与截止性测试表3样本清单核对，并补充未覆盖大额项目的选样说明。",
            "期末附近大额费用可能涉及截止性错报风险。",
        )


def analyze_sop_checklist(main_book: XLSXBook | None, uploaded_names: list[str], findings: list[dict[str, Any]]) -> None:
    if not main_book:
        return
    uploaded_text = " ".join(uploaded_names + main_book.sheets)
    for sheet, label in [
        (main_book.find_sheet("销售费用", "BKD"), "销售费用BKD"),
        (main_book.find_sheet("管理费用", "BKD"), "管理费用BKD"),
    ]:
        if not sheet:
            continue
        rows = main_book.rows(sheet)
        threshold = to_number(rows.get(23, {}).get(3)) or to_number(rows.get(22, {}).get(3)) or 0
        for row, cols in rows.items():
            if row < 30:
                continue
            xref = as_text(cols.get(14))
            account = as_text(cols.get(4))
            amount = abs(to_number(cols.get(5)) or 0)
            if not xref or amount < threshold:
                continue
            normalized = xref.strip("<>").split("_202")[0].replace("<", "").replace(">", "")
            if normalized and normalized not in uploaded_text:
                finding(
                    findings,
                    "Medium",
                    "Checklist",
                    f"{label}存在重大X-ref但未上传对应支持底稿",
                    main_book.name,
                    sheet,
                    f"N{row}",
                    f"{account} 本期金额={amount:,.2f}; X-ref={xref}",
                    "BKD中重大费用项目已索引至其他底稿，但本次上传材料中未识别到对应支持底稿。",
                    "请上传对应支持底稿，或在Notes中说明该X-ref在Canvas/项目文件中的具体位置。",
                    "X-ref不可追踪会影响底稿复核和审计证据完整性。",
                )
                break

    cutoff = main_book.find_sheet("截止性测试")
    if cutoff:
        for row, cols in main_book.rows(cutoff).items():
            if row < 27 or row > 37:
                continue
            voucher = as_text(cols.get(2))
            if voucher and voucher not in {"本期资产负债表日前记录的交易", "本期资产负债表日后记录的交易"}:
                ey_work = as_text(cols.get(11))
                if not ey_work:
                    finding(
                        findings,
                        "Medium",
                        "Checklist",
                        "截止性测试样本缺少EY Works结论",
                        main_book.name,
                        cutoff,
                        f"K{row}",
                        f"凭证编号={voucher}; EY Works为空",
                        "截止性测试样本未填写EY Works结论。",
                        "请补充Y/N结论；如为N，应在异常说明中记录差异原因、追加程序和最终结论。",
                        "测试结论缺失会影响截止性程序执行完整性。",
                    )
                    break


def model_status() -> dict[str, Any]:
    api_key = AI_CONFIG.get("api_key", "").strip()
    model = AI_CONFIG.get("model", "").strip()
    return {
        "enabled": bool(api_key),
        "model": model or "not configured",
        "mode": "llm_enhanced" if api_key else "rules_only",
        "message": "已配置大模型增强：规则负责定位，模型负责润色Review Notes和风险解释。"
        if api_key
        else "当前为规则引擎模式：可稳定定位风险点；配置OPENAI_API_KEY后可启用大模型增强。",
    }


def enhance_findings_with_model(payload: dict[str, Any]) -> dict[str, Any]:
    api_key = AI_CONFIG.get("api_key", "").strip()
    if not api_key or not payload.get("findings"):
        payload["ai"] = model_status()
        return payload

    model = AI_CONFIG.get("model", "gpt-4.1-mini").strip()
    endpoint = AI_CONFIG.get("endpoint", "https://api.openai.com/v1/responses").strip()
    compact_findings = [
        {
            "id": item["id"],
            "severity": item["severity"],
            "category": item["category"],
            "title": item["title"],
            "location": item["location"],
            "evidence": item["evidence"],
            "reviewNote": item["reviewNote"],
            "recommendation": item["recommendation"],
            "risk": item["risk"],
        }
        for item in payload["findings"]
    ]
    prompt = (
        "你是审计项目经理的底稿Review助手。请基于给定的规则发现生成更专业、简洁、可执行的中文Review Notes。"
        "不要新增不存在的单元格、文件、金额或事实；必须保留每条发现的id。"
        "请仅返回JSON对象，格式为{\"findings\":[...]}，数组中每个对象包含id、reviewNote、recommendation、risk三个字段。\n\n"
        + json.dumps(compact_findings, ensure_ascii=False)
    )
    request_body = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt}],
            }
        ],
        "text": {"format": {"type": "json_object"}},
    }
    request = urllib.request.Request(
        endpoint,
        data=json.dumps(request_body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=40) as response:
            response_payload = json.loads(response.read().decode("utf-8"))
        text = extract_response_text(response_payload)
        enhanced = json.loads(text)
        if isinstance(enhanced, dict):
            enhanced = enhanced.get("findings", enhanced.get("items", []))
        by_id = {item.get("id"): item for item in enhanced if isinstance(item, dict)}
        for item in payload["findings"]:
            update = by_id.get(item["id"])
            if not update:
                continue
            for key in ["reviewNote", "recommendation", "risk"]:
                if as_text(update.get(key)):
                    item[key] = as_text(update[key])
            item["aiEnhanced"] = True
        payload["ai"] = {
            "enabled": True,
            "model": model,
            "mode": "llm_enhanced",
            "message": "大模型增强已启用：Review Notes已基于规则发现进行审计口径润色。",
        }
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, KeyError, ValueError) as exc:
        payload["ai"] = {
            "enabled": False,
            "model": model,
            "mode": "rules_only",
            "message": f"大模型增强未完成，已回退到规则引擎结果：{exc}",
        }
    return payload


def extract_response_text(response_payload: dict[str, Any]) -> str:
    if "output_text" in response_payload:
        return as_text(response_payload["output_text"])
    chunks: list[str] = []
    for item in response_payload.get("output", []):
        for content in item.get("content", []):
            text = content.get("text")
            if text:
                chunks.append(text)
    return "\n".join(chunks)


def split_cell_ref(ref: str) -> tuple[int, int] | None:
    match = re.match(r"([A-Z]+)(\d+)", ref or "")
    if not match:
        return None
    col = 0
    for ch in match.group(1):
        col = col * 26 + ord(ch) - 64
    return int(match.group(2)), col


def first_cell(cell_range: str) -> str:
    match = re.search(r"[A-Z]+\d+", cell_range or "")
    return match.group(0) if match else "A1"


def safe_download_name(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", stem)[:80] or "workbook"


def annotation_text(item: dict[str, Any]) -> str:
    return (
        f"{item.get('id', '')} {item.get('title', '')}\n"
        f"风险等级：{item.get('severity', '')}\n"
        f"证据：{item.get('evidence', '')}\n"
        f"Review Note：{item.get('reviewNote', '')}\n"
        f"建议修改：{item.get('recommendation', '')}"
    )


def workbook_sheet_paths(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels: dict[str, str] = {}
        for rel in rel_root.findall("rel:Relationship", NS):
            target = rel.attrib["Target"]
            if target.startswith("/"):
                path = target.lstrip("/")
            elif target.startswith("xl/"):
                path = target
            else:
                path = f"xl/{target}"
            rels[rel.attrib["Id"]] = path
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        paths: dict[str, str] = {}
        for sheet in workbook.findall(".//a:sheet", NS):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            paths[sheet.attrib["name"]] = rels[rid]
        return paths


def next_part_number(names: set[str], prefix: str, suffix: str) -> int:
    pattern = re.compile(re.escape(prefix) + r"(\d+)" + re.escape(suffix) + r"$")
    current = 0
    for name in names:
        match = pattern.search(name)
        if match:
            current = max(current, int(match.group(1)))
    return current + 1


def next_rid(rel_root: ET.Element) -> str:
    current = 0
    for rel in rel_root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id", "")
        match = re.match(r"rId(\d+)$", rid)
        if match:
            current = max(current, int(match.group(1)))
    return f"rId{current + 1}"


def build_comments_xml(comments: list[dict[str, str]]) -> bytes:
    root = ET.Element(f"{{{NS['a']}}}comments")
    authors = ET.SubElement(root, f"{{{NS['a']}}}authors")
    ET.SubElement(authors, f"{{{NS['a']}}}author").text = "AI Review Assistant"
    comment_list = ET.SubElement(root, f"{{{NS['a']}}}commentList")
    for item in comments:
        comment = ET.SubElement(
            comment_list,
            f"{{{NS['a']}}}comment",
            {"ref": item["cell"], "authorId": "0", "shapeId": "0"},
        )
        text = ET.SubElement(comment, f"{{{NS['a']}}}text")
        run = ET.SubElement(text, f"{{{NS['a']}}}r")
        ET.SubElement(run, f"{{{NS['a']}}}t").text = item["text"]
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_vml(comments: list[dict[str, str]]) -> bytes:
    pieces = [
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:x="urn:schemas-microsoft-com:office:excel">',
        '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
        '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>',
    ]
    for idx, item in enumerate(comments, start=1025):
        parsed = split_cell_ref(item["cell"]) or (1, 1)
        row, col = parsed
        anchor = f"{col}, 15, {row}, 2, {col + 4}, 15, {row + 5}, 2"
        pieces.append(
            f'<v:shape id="_x0000_s{idx}" type="#_x0000_t202" '
            'style="position:absolute;margin-left:80pt;margin-top:5pt;width:220pt;height:110pt;z-index:1;visibility:hidden" '
            'fillcolor="#ffffe1" o:insetmode="auto">'
            '<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/>'
            '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>'
            '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
            f"<x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>"
            f"<x:Row>{row - 1}</x:Row><x:Column>{col - 1}</x:Column>"
            "</x:ClientData></v:shape>"
        )
    pieces.append("</xml>")
    return "".join(pieces).encode("utf-8")


def ensure_content_type(root: ET.Element, tag: str, attrs: dict[str, str]) -> None:
    for node in root.findall(f"{{{NS['ct']}}}{tag}"):
        if all(node.attrib.get(key) == value for key, value in attrs.items()):
            return
    ET.SubElement(root, f"{{{NS['ct']}}}{tag}", attrs)


def annotate_workbook(data: bytes, notes: list[dict[str, Any]]) -> bytes:
    if not notes:
        return data
    sheet_paths = workbook_sheet_paths(data)
    by_sheet: dict[str, list[dict[str, str]]] = {}
    for item in notes:
        location = item.get("location", {})
        sheet = location.get("sheet", "")
        if sheet not in sheet_paths:
            continue
        by_sheet.setdefault(sheet, []).append(
            {"cell": first_cell(location.get("cell", "")), "text": annotation_text(item)}
        )
    if not by_sheet:
        return data

    source = zipfile.ZipFile(io.BytesIO(data))
    names = set(source.namelist())
    overrides: dict[str, bytes] = {}

    content_root = ET.fromstring(source.read("[Content_Types].xml"))
    ensure_content_type(
        content_root,
        "Default",
        {"Extension": "vml", "ContentType": "application/vnd.openxmlformats-officedocument.vmlDrawing"},
    )

    comment_no = next_part_number(names, "xl/comments", ".xml")
    vml_no = next_part_number(names, "xl/drawings/vmlDrawing", ".vml")

    for sheet_name, comments in by_sheet.items():
        sheet_path = sheet_paths[sheet_name]
        rels_path = str(Path(sheet_path).parent / "_rels" / (Path(sheet_path).name + ".rels")).replace("\\", "/")
        if rels_path in names:
            rel_root = ET.fromstring(source.read(rels_path))
        else:
            rel_root = ET.Element(f"{{{NS['rel']}}}Relationships")

        comments_path = f"xl/comments{comment_no}.xml"
        vml_path = f"xl/drawings/vmlDrawing{vml_no}.vml"
        comment_no += 1
        vml_no += 1

        comments_rid = next_rid(rel_root)
        ET.SubElement(
            rel_root,
            f"{{{NS['rel']}}}Relationship",
            {
                "Id": comments_rid,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "Target": f"../{Path(comments_path).name}",
            },
        )
        vml_rid = next_rid(rel_root)
        ET.SubElement(
            rel_root,
            f"{{{NS['rel']}}}Relationship",
            {
                "Id": vml_rid,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                "Target": f"../drawings/{Path(vml_path).name}",
            },
        )

        sheet_xml = source.read(sheet_path)
        if b"legacyDrawing" not in sheet_xml:
            insert = f'<legacyDrawing r:id="{vml_rid}"/>'.encode("utf-8")
            sheet_xml = sheet_xml.replace(b"</worksheet>", insert + b"</worksheet>", 1)
            overrides[sheet_path] = sheet_xml
        overrides[rels_path] = ET.tostring(rel_root, encoding="utf-8", xml_declaration=True)
        overrides[comments_path] = build_comments_xml(comments)
        overrides[vml_path] = build_vml(comments)
        ensure_content_type(
            content_root,
            "Override",
            {
                "PartName": f"/{comments_path}",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            },
        )

    overrides["[Content_Types].xml"] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename in overrides:
                continue
            target.writestr(info, source.read(info.filename))
        for name, content in overrides.items():
            target.writestr(name, content)
    source.close()
    return output.getvalue()


def package_annotated_workbooks(files: list[tuple[str, bytes]], payload: dict[str, Any]) -> str | None:
    findings = payload.get("findings", [])
    if not findings:
        return None
    by_file: dict[str, list[dict[str, Any]]] = {}
    for item in findings:
        file_name = item.get("location", {}).get("file")
        if file_name:
            by_file.setdefault(file_name, []).append(item)
    if not by_file:
        return None

    run_id = f"{int(time.time() * 1000)}_{uuid.uuid4().hex[:8]}"
    zip_name = f"annotated_review_workpapers_{run_id}.zip"
    output = io.BytesIO()
    wrote = False
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as bundle:
        for name, data in files:
            if not name.lower().endswith(".xlsx") or name not in by_file:
                continue
            annotated = annotate_workbook(data, by_file[name])
            out_name = f"{safe_download_name(name)}_AI批注版.xlsx"
            bundle.writestr(out_name, annotated)
            wrote = True
        bundle.writestr(
            "Review Notes.json",
            json.dumps(findings, ensure_ascii=False, indent=2).encode("utf-8"),
        )
    if not wrote:
        return None
    DOWNLOADS[zip_name] = output.getvalue()
    return f"/downloads/{zip_name}"


def analyze_files(
    files: list[tuple[str, bytes]],
    include_support_checks: bool = True,
    explicit_tb_file: tuple[str, bytes] | None = None,
    explicit_je_file: tuple[str, bytes] | None = None,
    annotation_files: list[tuple[str, bytes]] | None = None,
) -> dict[str, Any]:
    books: list[XLSXBook] = []
    docs: list[dict[str, Any]] = []
    supporting: list[dict[str, Any]] = []
    tb_file: tuple[str, bytes] | None = explicit_tb_file
    je_file: tuple[str, bytes] | None = explicit_je_file
    file_errors: list[dict[str, str]] = []
    for name, data in files:
        lower = name.lower()
        try:
            if lower.endswith(".xlsx"):
                if "序时账" in name or "科目余额表" in name:
                    supporting.append({"name": name, "kind": "supporting", "size": len(data)})
                    if "科目余额表" in name and tb_file is None:
                        tb_file = (name, data)
                    if "序时账" in name and je_file is None:
                        je_file = (name, data)
                    continue
                books.append(XLSXBook(name, data))
            elif lower.endswith(".docx"):
                docs.append(parse_docx(name, data))
        except Exception as exc:
            file_errors.append({"name": name, "error": str(exc)})
    if explicit_tb_file:
        supporting.append({"name": explicit_tb_file[0], "kind": "supporting", "size": len(explicit_tb_file[1])})
    if explicit_je_file:
        supporting.append({"name": explicit_je_file[0], "kind": "supporting", "size": len(explicit_je_file[1])})

    findings: list[dict[str, Any]] = []
    main_book, tod_book = identify_books(books)
    context: dict[str, Any] = {}
    if main_book:
        context = analyze_main_book(main_book, tod_book, findings)
    else:
        finding(
            findings,
            "High",
            "材料缺失",
            "未识别到U_exp主底稿",
            "",
            "",
            "",
            "上传文件中未找到包含 Uexp.00 Lead 的工作簿。",
            "本次分析未识别到销售费用/管理费用主底稿。",
            "请上传U_exp SWP VC&VD主底稿后重新分析。",
            "缺少主底稿时无法完成程序完整性检查。",
        )
    if tod_book:
        analyze_tod_book(tod_book, context, findings)
    analyze_sop_checklist(main_book, [name for name, _data in files], findings)
    if include_support_checks and tb_file:
        analyze_tb_support(tb_file[0], tb_file[1], main_book, findings)
    if include_support_checks and je_file:
        analyze_je_support(je_file[0], je_file[1], context, findings, tod_book, main_book)

    summary = {
        "total": len(findings),
        "high": sum(1 for item in findings if item["severity"] == "High"),
        "medium": sum(1 for item in findings if item["severity"] == "Medium"),
        "low": sum(1 for item in findings if item["severity"] == "Low"),
    }
    payload = {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "stats": summary,
        "files": [
            {"name": book.name, "kind": "workbook", "sheets": book.sheets}
            for book in books
        ]
        + docs
        + supporting,
        "errors": file_errors,
        "findings": findings,
        "options": {"supportChecks": include_support_checks},
    }
    payload = enhance_findings_with_model(payload)
    annotated_url = package_annotated_workbooks(annotation_files or files, payload)
    payload["annotated"] = {
        "available": bool(annotated_url),
        "url": annotated_url,
        "message": "已生成批注版底稿，可下载后在Excel中查看问题批注。"
        if annotated_url
        else "未生成批注版底稿：本次发现未定位到可回写的Excel单元格。",
    }
    return payload


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        sys.stdout.write("%s - %s\n" % (self.address_string(), format % args))

    def do_GET(self) -> None:
        raw_path = unquote(self.path.split("?", 1)[0])
        if raw_path == "/":
            self._send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if raw_path.startswith("/downloads/"):
            download_name = raw_path.removeprefix("/downloads/")
            if download_name in DOWNLOADS:
                self._send_bytes(DOWNLOADS[download_name], "application/zip", download_name)
                return
            self.send_error(404)
            return
        requested = (STATIC_DIR / raw_path.lstrip("/")).resolve()
        if STATIC_DIR.resolve() in requested.parents and requested.exists():
            content_type = "text/plain; charset=utf-8"
            if requested.suffix == ".css":
                content_type = "text/css; charset=utf-8"
            elif requested.suffix == ".js":
                content_type = "application/javascript; charset=utf-8"
            self._send_file(requested, content_type)
            return
        self.send_error(404)

    def do_POST(self) -> None:
        try:
            parsed = urlparse(self.path)
            path = parsed.path
            query = parse_qs(parsed.query)
            if path == "/api/config":
                length = int(self.headers.get("Content-Length", "0") or "0")
                raw = self.rfile.read(length) if length else b"{}"
                config = json.loads(raw.decode("utf-8") or "{}")
                api_key = as_text(config.get("apiKey"))
                model = as_text(config.get("model")) or "gpt-4.1-mini"
                endpoint = as_text(config.get("endpoint")) or "https://api.openai.com/v1/responses"
                if api_key:
                    AI_CONFIG["api_key"] = api_key
                if config.get("clearApiKey"):
                    AI_CONFIG["api_key"] = ""
                AI_CONFIG["model"] = model
                AI_CONFIG["endpoint"] = endpoint
                self._send_json({"ok": True, "ai": model_status(), "endpoint": AI_CONFIG["endpoint"]})
                return
            if path == "/api/analyze-sample":
                self._send_json({"error": "样例分析已关闭。请在页面中自行选择底稿、TB和JE后分析。"}, status=410)
                return
            if path == "/api/analyze":
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type", "")},
                )
                workpapers = self._form_files(form, "workpapers")
                programs = self._form_files(form, "programs")
                fallback = self._form_files(form, "files")
                tb_uploads = self._form_files(form, "tb")
                je_uploads = self._form_files(form, "je")
                uploads = workpapers + programs + fallback
                if not uploads:
                    self._send_json({"error": "没有收到审计底稿，请至少在“底稿”区域选择一个Excel底稿。"}, status=400)
                    return
                support_value = form.getfirst("supportChecks", "1")
                self._send_json(
                    analyze_files(
                        uploads,
                        include_support_checks=support_value != "0",
                        explicit_tb_file=tb_uploads[0] if tb_uploads else None,
                        explicit_je_file=je_uploads[0] if je_uploads else None,
                        annotation_files=workpapers or uploads,
                    )
                )
                return
            self.send_error(404)
        except Exception as exc:
            self._send_json({"error": f"分析失败：{exc.__class__.__name__}: {exc}"}, status=500)

    def _form_files(self, form: cgi.FieldStorage, field_name: str) -> list[tuple[str, bytes]]:
        if field_name not in form:
            return []
        fields = form[field_name]
        if not isinstance(fields, list):
            fields = [fields]
        uploads: list[tuple[str, bytes]] = []
        for field in fields:
            if field.filename:
                uploads.append((Path(field.filename).name, field.file.read()))
        return uploads

    def do_GET(self) -> None:
        raw_path = unquote(self.path.split("?", 1)[0])
        if raw_path == "/api/health":
            self._send_json({"ok": True, "service": "AI Audit Review Assistant"})
            return
        if raw_path == "/api/config":
            self._send_json({"ai": model_status(), "endpoint": AI_CONFIG["endpoint"]})
            return
        if raw_path == "/":
            self._send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if raw_path.startswith("/downloads/"):
            download_name = raw_path.removeprefix("/downloads/")
            if download_name in DOWNLOADS:
                self._send_bytes(DOWNLOADS[download_name], "application/zip", download_name)
                return
            self.send_error(404)
            return
        requested = (STATIC_DIR / raw_path.lstrip("/")).resolve()
        if STATIC_DIR.resolve() in requested.parents and requested.exists():
            content_type = "text/plain; charset=utf-8"
            if requested.suffix == ".css":
                content_type = "text/css; charset=utf-8"
            elif requested.suffix == ".js":
                content_type = "application/javascript; charset=utf-8"
            self._send_file(requested, content_type)
            return
        self.send_error(404)

    def _send_file(self, path: Path, content_type: str) -> None:
        data = path.read_bytes()
        self._send_bytes(data, content_type, path.name)

    def _send_bytes(self, data: bytes, content_type: str, filename: str | None = None) -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        if filename and content_type == "application/zip":
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    port = 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Audit Review Assistant running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
