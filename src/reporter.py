"""
审计底稿质检Agent - 结果输出模块
负责将质检结果输出为Excel批注和汇总报告
"""

import pandas as pd
import re
from dataclasses import replace
from collections import OrderedDict
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from typing import List, Optional
from datetime import datetime

from .models import AuditWorkbook, CellLocation, CheckResult, QualityCheckReport, Severity, CheckCategory, RuleExecutionRecord, RuleStatus
from .check_logic import build_execution_summary


class QualityReporter:
    """质检结果输出器"""

    MAX_INLINE_EVIDENCE_LOCATIONS = 5
    FONT_NAME = "等线"
    EY_YELLOW = "FFE600"
    EY_BLACK = "111111"
    EY_CHARCOAL = "1F1F1F"
    EY_GRAPHITE = "333333"
    EY_LINE = "D9D9D9"
    EY_LIGHT_BG = "F5F5F3"
    HEADER_FILL = PatternFill("solid", fgColor=EY_BLACK)
    TITLE_FILL = PatternFill("solid", fgColor=EY_BLACK)
    SUBTITLE_FILL = PatternFill("solid", fgColor="F2F2EF")
    METRIC_LABEL_FILL = PatternFill("solid", fgColor=EY_CHARCOAL)
    METRIC_VALUE_FILL = PatternFill("solid", fgColor="F2F2EF")
    SECTION_FILL = PatternFill("solid", fgColor=EY_CHARCOAL)
    FAILED_FILL = PatternFill("solid", fgColor="FCE8E6")
    PASSED_FILL = PatternFill("solid", fgColor="E6F3ED")
    SKIPPED_FILL = PatternFill("solid", fgColor="EDEDED")
    AI_DISABLED_FILL = PatternFill("solid", fgColor="E8EEF4")
    REVIEW_FILL = PatternFill("solid", fgColor="FFF4B8")
    MEDIUM_FILL = PatternFill("solid", fgColor="FFF0D8")
    LOW_FILL = PatternFill("solid", fgColor="E7F2EC")
    DETAIL_LEFT_FILL = PatternFill("solid", fgColor="F2F2F2")
    DETAIL_SKIPPED_FILL = PatternFill("solid", fgColor="FAFAFA")
    YELLOW_FILL = PatternFill("solid", fgColor=EY_YELLOW)
    WHITE_FONT = Font(name=FONT_NAME, color="FFFFFF", bold=True)
    TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True, color="FFFFFF")
    SUBTITLE_FONT = Font(name=FONT_NAME, size=10, color="595959")
    BODY_FONT = Font(name=FONT_NAME, size=10, color="333333")
    METRIC_LABEL_FONT = Font(name=FONT_NAME, size=10, color="FFFFFF", bold=True)
    METRIC_VALUE_FONT = Font(name=FONT_NAME, size=11, color=EY_BLACK, bold=True)
    LINK_FONT = Font(name=FONT_NAME, color="0563C1", underline="single")
    THIN_BORDER = Border(
        left=Side(style="thin", color=EY_LINE),
        right=Side(style="thin", color=EY_LINE),
        top=Side(style="thin", color=EY_LINE),
        bottom=Side(style="thin", color=EY_LINE),
    )
    FREEZE_BORDER = Border(top=Side(style="medium", color="777777"))

    def __init__(self):
        """初始化输出器"""
        pass

    def generate_report(self, workbook: AuditWorkbook, results: List[CheckResult],
                        output_path: str, execution_records: List[RuleExecutionRecord] = None) -> str:
        """
        生成质检报告

        Args:
            workbook: 工作簿对象
            results: 检查结果列表
            output_path: 输出文件路径

        Returns:
            str: 输出文件路径
        """
        # 原始结果用于逐单元格批注；聚合结果用于报告展示和问题数量统计。
        display_results = self._aggregate_results(results)
        report = QualityCheckReport(workbook=workbook, results=display_results)

        # 加载原始Excel文件
        wb = load_workbook(workbook.file_path)

        # 1. 在出现问题的单元格添加批注
        self._add_comments(wb, results)

        # 2. 在最前面添加汇总报告Sheet
        self._add_summary_sheet(wb, display_results, report.get_summary(), execution_records)

        # 3. 添加核对逻辑Sheet（第二个Sheet）
        self._add_check_logic_sheet(wb, display_results, execution_records)

        self._ensure_report_sheet_order(wb)
        self._apply_report_workbook_view(wb)

        # 保存文件
        wb.save(output_path)
        return output_path

    @staticmethod
    def _normalize_issue_text(text: Optional[str]) -> str:
        """去除单元格、编号、金额等动态内容，保留问题语义用于同类归组。"""
        normalized = re.sub(r"\s+", "", str(text or ""))
        normalized = re.sub(r"['“‘][^'”’]{1,80}['”’]", "<对象>", normalized)
        normalized = re.sub(r"累计折旧|减值准备|原值|净值", "<报表项目>", normalized)
        normalized = re.sub(r"上期末|期初|期末", "<期间>", normalized)
        normalized = re.sub(r"(?<![A-Za-z])\$?[A-Z]{1,3}\$?\d+", "<单元格>", normalized)
        normalized = re.sub(r"(?<![A-Za-z])[-+]?\d[\d,]*(?:\.\d+)?%?", "<数值>", normalized)
        normalized = re.sub(r"第<数值>行", "第<行>行", normalized)
        return normalized

    def _issue_group_key(self, result: CheckResult) -> tuple:
        if result.issue_group:
            issue_type = result.issue_group
        else:
            issue_type = self._normalize_issue_text(result.message)
        return (
            result.rule_id,
            result.rule_name,
            result.category,
            result.severity,
            issue_type,
            self._normalize_issue_text(result.expected),
        )

    @staticmethod
    def _unique_text(values) -> List[str]:
        seen = set()
        output = []
        for value in values:
            text = str(value or "").strip()
            if text and text not in seen:
                seen.add(text)
                output.append(text)
        return output

    @staticmethod
    def _location_label(result: CheckResult) -> str:
        if not result.show_location:
            return ""
        return f"{result.location.sheet}!{result.location.to_excel_address()}"

    @staticmethod
    def _common_issue_summary(group: List[CheckResult]) -> str:
        """将带有不同对象、金额的同类明细概括为简短共性描述。"""
        first = str(group[0].message or "").strip().rstrip("。；")
        compact = re.sub(r"\s+", "", first)

        patterns = [
            (
                r"FAList中资产类别.+折旧年限.+不在K\.03\.3列示的.+折旧年限范围.+内",
                "FA List中部分资产折旧年限不在K.03.3政策范围内",
            ),
            (
                r"选样输出中的样本.+未在.+实际测试样本中找到",
                "选样输出中的部分样本未在实际测试样本中找到",
            ),
            (
                r"K\.01BKD.+与TB存在差异",
                "K.01 BKD与TB存在差异",
            ),
            (
                r"K\.01.+表2checkwith表1.+区域差异",
                "K.01“表2 check with 表1”区域存在差异",
            ),
            (
                r".+借贷不平衡.+调整明细金额合计",
                "调整明细存在借贷不平衡",
            ),
        ]
        for pattern, summary in patterns:
            if re.search(pattern, compact):
                return summary

        # 优先使用规则名称作为稳定的共性描述，避免把具体对象、金额或地址重复写入。
        rule_name = str(group[0].rule_name or "").strip()
        return f"{rule_name}存在同类异常" if rule_name else "存在同类异常"

    def _aggregate_results(self, results: List[CheckResult]) -> List[CheckResult]:
        """同一检查点的同类问题合并展示，保留全部问题位置。"""
        grouped = OrderedDict()
        for result in results:
            grouped.setdefault(self._issue_group_key(result), []).append(result)

        aggregated = []
        for group in grouped.values():
            first = group[0]
            locations = []
            seen_locations = set()
            for result in group:
                if not result.show_location:
                    continue
                location_key = (
                    result.location.sheet,
                    result.location.row,
                    result.location.col,
                )
                if location_key not in seen_locations:
                    seen_locations.add(location_key)
                    locations.append(result.location)

            if len(group) == 1:
                aggregated.append(replace(first, related_locations=locations))
                continue

            messages = self._unique_text(result.message for result in group)
            if len(messages) == 1:
                message = f"{messages[0].rstrip('。；')}，共{len(locations) or len(group)}处。"
            else:
                common_summary = self._common_issue_summary(group)
                message = f"{common_summary}，共{len(locations) or len(group)}处。"

            actual_parts = []
            for result in group:
                label = self._location_label(result)
                actual = str(result.actual or "").strip()
                if actual:
                    actual_parts.append(f"{label}={actual}" if label else actual)

            expected_values = self._unique_text(result.expected for result in group)
            suggestion_values = self._unique_text(result.ai_suggestion for result in group)
            aggregated.append(replace(
                first,
                message=message,
                expected="；".join(expected_values),
                actual="；".join(actual_parts),
                ai_suggestion="；".join(suggestion_values),
                related_locations=locations,
            ))
        return aggregated

    def _ensure_report_sheet_order(self, wb):
        """确保报告相关Sheet固定显示在工作簿最前面。"""
        report_sheet_names = ["质检报告", "核对逻辑"]
        report_sheets = [wb[name] for name in report_sheet_names if name in wb.sheetnames]
        other_sheets = [sheet for sheet in wb._sheets if sheet not in report_sheets]
        wb._sheets = report_sheets + other_sheets

    def _apply_report_workbook_view(self, wb):
        """让用户打开文件时默认看到质检报告页，原底稿Sheet格式保持不动。"""
        if "质检报告" in wb.sheetnames:
            wb.active = wb.sheetnames.index("质检报告")

    def _style_report_title(self, ws, max_col: int, summary_last_row: int = 2):
        ws["A1"].font = self.TITLE_FONT
        ws["A1"].fill = self.TITLE_FILL
        ws["A1"].alignment = Alignment(vertical="center", horizontal="left")
        ws.row_dimensions[1].height = 30
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.TITLE_FILL
            cell.border = self.THIN_BORDER
            if col != 1:
                cell.font = self.TITLE_FONT
        for row in range(2, summary_last_row + 1):
            ws.row_dimensions[row].height = 22
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = self.SUBTITLE_FONT if row == 2 else self.BODY_FONT
                cell.fill = PatternFill(fill_type=None)
                cell.border = self.THIN_BORDER
                cell.alignment = Alignment(vertical="center", wrap_text=False)

    def _style_table_header(self, ws, header_row: int, max_col: int):
        ws.row_dimensions[header_row].height = 26
        header_border = Border(
            left=Side(style="thin", color=self.EY_LINE),
            right=Side(style="thin", color=self.EY_LINE),
            top=Side(style="thin", color=self.EY_LINE),
            bottom=Side(style="medium", color=self.EY_YELLOW),
        )
        for col in range(1, max_col + 1):
            cell = ws.cell(row=header_row, column=col)
            cell.font = self.WHITE_FONT
            cell.fill = self.HEADER_FILL
            cell.border = header_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _style_table_body(self, ws, start_row: int, end_row: int, max_col: int):
        if end_row < start_row:
            return
        for row in range(start_row, end_row + 1):
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = self.LINK_FONT if cell.hyperlink else self.BODY_FONT
                cell.fill = PatternFill(fill_type=None)
                cell.border = self.THIN_BORDER
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for col in (1,):
                ws.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            self._fit_row_height(ws, row, max_col)

    @staticmethod
    def _display_text_width(text: str) -> float:
        width = 0.0
        for char in str(text or ""):
            if char in "\r\n":
                continue
            width += 2.0 if ord(char) > 127 else 1.0
        return width

    def _fit_row_height(self, ws, row: int, max_col: int, min_height: int = 18, max_height: int = 120):
        max_lines = 1
        for col in range(1, max_col + 1):
            value = ws.cell(row=row, column=col).value
            if value in (None, ""):
                continue
            column_letter = get_column_letter(col)
            column_width = ws.column_dimensions[column_letter].width or 10
            usable_width = max(column_width, 8)
            cell_lines = 0
            for part in str(value).splitlines() or [""]:
                text_width = self._display_text_width(part)
                cell_lines += max(1, int((text_width + usable_width - 1) // usable_width))
            max_lines = max(max_lines, cell_lines)
        ws.row_dimensions[row].height = min(max_height, max(min_height, 18 + (max_lines - 1) * 15))

    def _style_metric_pairs(self, ws, rows: range, max_col: int):
        for row in rows:
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value in (None, ""):
                    continue
                cell.border = self.THIN_BORDER
                cell.font = self.METRIC_LABEL_FONT if col % 2 == 1 else self.METRIC_VALUE_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
                cell.fill = self.METRIC_LABEL_FILL if col % 2 == 1 else self.METRIC_VALUE_FILL

    def _summary_status_fill(self, status: str):
        return {
            RuleStatus.FAILED.value: self.FAILED_FILL,
            RuleStatus.PASSED.value: self.PASSED_FILL,
            RuleStatus.SKIPPED.value: self.DETAIL_SKIPPED_FILL,
            RuleStatus.AI_DISABLED.value: self.AI_DISABLED_FILL,
        }.get(status)

    def _detail_right_fill(self, status: str, severity: str = ""):
        if status == RuleStatus.FAILED.value and severity == Severity.REVIEW.value:
            return self.REVIEW_FILL
        if status == RuleStatus.FAILED.value:
            return self.FAILED_FILL
        return self._summary_status_fill(status)

    def _severity_fill(self, severity: str):
        return {
            Severity.HIGH.value: self.FAILED_FILL,
            Severity.MEDIUM.value: self.MEDIUM_FILL,
            Severity.LOW.value: self.LOW_FILL,
            Severity.REVIEW.value: self.REVIEW_FILL,
        }.get(severity)

    def _apply_status_cell_fill(self, ws, row: int, status_col: int):
        fill = self._summary_status_fill(ws.cell(row=row, column=status_col).value)
        if fill:
            cell = ws.cell(row=row, column=status_col)
            cell.fill = fill
            cell.font = Font(name=self.FONT_NAME, size=10, color=self.EY_BLACK, bold=True)

    def _apply_detail_band_fill(self, ws, start_row: int, end_row: int, max_col: int,
                                status_col: int, severity_col: Optional[int] = None,
                                left_cols: int = 4):
        """明细行左侧固定浅灰，右侧按状态/严重程度着色。"""
        if end_row < start_row:
            return
        for row in range(start_row, end_row + 1):
            status = ws.cell(row=row, column=status_col).value
            severity = ws.cell(row=row, column=severity_col).value if severity_col else ""
            self._apply_detail_band_fill_for_status(
                ws, row, max_col, status, severity, left_cols=left_cols
            )
            ws.cell(row=row, column=status_col).font = Font(
                name=self.FONT_NAME, size=10, color=self.EY_BLACK, bold=True
            )
            if severity_col:
                ws.cell(row=row, column=severity_col).font = Font(
                    name=self.FONT_NAME, size=10, color=self.EY_BLACK, bold=True
                )

    def _apply_detail_band_fill_for_status(self, ws, row: int, max_col: int,
                                           status: str, severity: str = "",
                                           left_cols: int = 4):
        right_fill = self._detail_right_fill(status, severity)
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = self.DETAIL_LEFT_FILL if col <= left_cols else (right_fill or PatternFill(fill_type=None))

    def _style_summary_sheet(self, ws, header_row: int, last_row: int):
        max_col = 10
        ws["A5"] = ""
        self._style_report_title(ws, max_col, header_row - 2)
        metric_rows = range(3, 5) if header_row == 7 else range(3, header_row - 1)
        self._style_metric_pairs(ws, metric_rows, max_col)
        self._style_table_header(ws, header_row, max_col)
        self._style_table_body(ws, header_row + 1, last_row, max_col)
        self._apply_detail_band_fill(ws, header_row + 1, last_row, max_col, status_col=5, severity_col=6)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
        ws.auto_filter.ref = f"A{header_row}:J{max(last_row, header_row)}"
        ws.sheet_properties.tabColor = self.EY_YELLOW

    def _style_logic_sheet(self, ws, header_row: int, last_row: int):
        max_col = 8
        self._style_report_title(ws, max_col, 2)
        self._style_table_header(ws, header_row, max_col)
        self._style_table_body(ws, header_row + 1, last_row, max_col)
        self._apply_detail_band_fill(ws, header_row + 1, last_row, max_col, status_col=4)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
        ws.auto_filter.ref = f"A{header_row}:H{max(last_row, header_row)}"
        ws.sheet_properties.tabColor = self.EY_GRAPHITE

    def _add_comments(self, wb, results: List[CheckResult]):
        """在问题单元格添加批注"""
        from openpyxl.cell.cell import MergedCell
        for result in results:
            if not result.add_comment:
                continue
            sheet_name = result.location.sheet
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                cell = ws.cell(
                    row=result.location.row,
                    column=result.location.col
                )

                # 合并单元格将批注写到合并区域左上角。
                if isinstance(cell, MergedCell):
                    for merged_range in ws.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            cell = ws.cell(
                                row=merged_range.min_row,
                                column=merged_range.min_col,
                            )
                            break

                # 创建批注
                comment_text = f"[{result.rule_id}] {result.message}"
                if result.expected:
                    comment_text += f"\n期望: {result.expected}"
                if result.actual:
                    comment_text += f"\n实际: {result.actual}"

                if cell.comment and comment_text not in cell.comment.text:
                    comment_text = f"{cell.comment.text}\n\n{comment_text}"
                comment = Comment(
                    comment_text,
                    f"质检系统 - {result.severity.value}级问题"
                )
                comment.width = 300
                comment.height = 100

                # 添加批注到单元格
                cell.comment = comment

    def _add_summary_sheet(self, wb, results: List[CheckResult], summary: dict,
                           execution_records: List[RuleExecutionRecord] = None):
        """在最前面添加汇总报告Sheet"""
        # 创建新的汇总Sheet
        summary_sheet_name = "质检报告"
        if summary_sheet_name in wb.sheetnames:
            del wb[summary_sheet_name]

        ws = wb.create_sheet(summary_sheet_name, 0)  # 插入到最前面

        # 写入标题
        ws['A1'] = "审计底稿质检报告"
        ws['A1'].font = Font(size=16, bold=True)
        ws['A2'] = f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        if execution_records:
            passed = sum(1 for er in execution_records if er.status == RuleStatus.PASSED)
            failed = sum(1 for er in execution_records if er.status == RuleStatus.FAILED)
            skipped = sum(1 for er in execution_records if er.status == RuleStatus.SKIPPED)
            ai_disabled = sum(1 for er in execution_records if er.status == RuleStatus.AI_DISABLED)
            ws['A3'], ws['B3'] = "规则总数", len(execution_records)
            ws['C3'], ws['D3'] = "发现问题", failed
            ws['E3'], ws['F3'] = "检查通过", passed
            ws['G3'], ws['H3'] = "未检查", skipped
            ws['I3'], ws['J3'] = "未调用AI", ai_disabled
            ws['A4'], ws['B4'] = "问题总数", summary['total']
            ws['C4'], ws['D4'] = "高严重度", summary['high']
            ws['E4'], ws['F4'] = "中严重度", summary['medium']
            ws['G4'], ws['H4'] = "低严重度", summary['low']
            ws['I4'], ws['J4'] = "需review", summary['review']
            header_row = 7
        else:
            ws['A3'], ws['B3'] = "问题总数", summary['total']
            ws['C3'], ws['D3'] = "高严重度", summary['high']
            ws['E3'], ws['F3'] = "中严重度", summary['medium']
            ws['G3'], ws['H3'] = "低严重度", summary['low']
            ws['I3'], ws['J3'] = "需review", summary['review']
            header_row = 6

        # 写入表头
        headers = [
            "序号", "规则ID", "规则名称", "检查类别", "执行状态", "严重程度",
            "所在Sheet", "单元格地址", "问题描述", "复核建议"
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col, value=header)
            cell.font = Font(bold=True)

        current_row = header_row + 1
        row_index = 1

        def write_row(rule_id, rule_name, category, status, severity="",
                      sheet="", address="", message="", ai_suggestion="",
                      hyperlink_sheet=None, hyperlink_address=None):
            nonlocal current_row, row_index
            ws.cell(row=current_row, column=1, value=row_index)
            ws.cell(row=current_row, column=2, value=rule_id)
            ws.cell(row=current_row, column=3, value=rule_name)
            ws.cell(row=current_row, column=4, value=category)
            ws.cell(row=current_row, column=5, value=status)
            ws.cell(row=current_row, column=6, value=severity)
            ws.cell(row=current_row, column=7, value=sheet)
            # 单元格地址列（加超链接，点击跳转到底稿对应单元格）
            addr_cell = ws.cell(row=current_row, column=8, value=address)
            link_sheet = hyperlink_sheet if hyperlink_sheet is not None else sheet
            link_address = hyperlink_address if hyperlink_address is not None else address
            if link_sheet and link_address:
                try:
                    # 用safe_sheet_name处理含特殊字符的sheet名
                    hyperlink_sheet, hyperlink_address = self._first_hyperlink_target(
                        link_sheet, link_address
                    )
                    safe_sheet = hyperlink_sheet.replace("'", "''")
                    addr_cell.hyperlink = f"#'{safe_sheet}'!{hyperlink_address}"
                    addr_cell.font = self.LINK_FONT
                except Exception:
                    pass
            ws.cell(row=current_row, column=9, value=message)
            ws.cell(row=current_row, column=10, value=ai_suggestion)
            current_row += 1
            row_index += 1

        if execution_records:
            # 写入所有规则状态；有问题的规则按问题明细展开。
            def record_sort_key(record: RuleExecutionRecord):
                if record.status == RuleStatus.FAILED:
                    has_review_issue = any(
                        issue.severity == Severity.REVIEW for issue in record.issues
                    )
                    return 1 if has_review_issue else 0
                if record.status == RuleStatus.PASSED:
                    return 2
                if record.status == RuleStatus.SKIPPED:
                    return 3
                if record.status == RuleStatus.AI_DISABLED:
                    return 4
                return 99

            ordered_records = sorted(
                execution_records,
                key=record_sort_key,
            )
            for er in ordered_records:
                if er.issues:
                    display_issues = sorted(
                        self._aggregate_results(er.issues),
                        key=lambda issue: 1 if issue.severity == Severity.REVIEW else 0,
                    )
                    for result in display_issues:
                        sheet, address = self._display_location(result)
                        write_row(
                            result.rule_id,
                            result.rule_name,
                            result.category.value,
                            er.status.value,
                            result.severity.value,
                            sheet,
                            address,
                            result.message,
                            result.ai_suggestion or "",
                        )
                else:
                    if er.status == RuleStatus.PASSED:
                        message = "未发现问题"
                        sheet, address, hyperlink_sheet, hyperlink_address = self._display_record_evidence(er)
                    else:
                        message = er.skipped_reason or er.status.value
                        sheet, address, hyperlink_sheet, hyperlink_address = "", "", "", ""
                    write_row(
                        er.rule_id,
                        er.rule_name,
                        er.category.value,
                        er.status.value,
                        sheet=sheet,
                        address=address,
                        message=message,
                        hyperlink_sheet=hyperlink_sheet,
                        hyperlink_address=hyperlink_address,
                    )
        else:
            # 兼容未传入执行记录的旧调用方式：只输出问题明细。
            for result in results:
                sheet, address = self._display_location(result)
                write_row(
                    result.rule_id,
                    result.rule_name,
                    result.category.value,
                    "发现问题",
                    result.severity.value,
                    sheet,
                    address,
                    result.message,
                    result.ai_suggestion or "",
                )

        # 调整列宽
        column_widths = [10, 12, 25, 15, 12, 10, 25, 15, 45, 30]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        self._style_summary_sheet(ws, header_row, current_row - 1)

    @staticmethod
    def _display_location(result: CheckResult) -> tuple:
        if not result.show_location:
            return "", ""
        return QualityReporter._display_locations(result.related_locations or [result.location])

    @staticmethod
    def _display_locations(locations: List[CellLocation]) -> tuple:
        sheets = []
        for location in locations:
            if location.sheet not in sheets:
                sheets.append(location.sheet)
        if len(sheets) == 1:
            return sheets[0], "、".join(
                location.to_excel_address() for location in locations
            )
        return "多个Sheet", "；".join(
            f"{location.sheet}!{location.to_excel_address()}"
            for location in locations
        )

    @staticmethod
    def _display_record_evidence(record: RuleExecutionRecord) -> tuple:
        """返回执行记录证据的展示Sheet、地址、可跳转Sheet、可跳转地址。"""
        if record.evidence_locations:
            sheet, address = QualityReporter._display_evidence_locations(
                record.evidence_locations,
                record.evidence_location_count,
            )
            hyperlink_sheet, hyperlink_address = QualityReporter._display_locations([
                record.evidence_locations[0]
            ])
            display_address = f"{address}；{record.evidence_text}" if record.evidence_text else address
            return sheet, display_address, hyperlink_sheet, hyperlink_address
        if record.evidence_text:
            return "核对范围", record.evidence_text, "", ""
        return "", "", "", ""

    @staticmethod
    def _display_evidence_locations(locations: List[CellLocation],
                                    total_count: int = 0) -> tuple:
        """少量位置直接展示；大量位置按Sheet聚合，避免报告刷屏。"""
        if len(locations) <= QualityReporter.MAX_INLINE_EVIDENCE_LOCATIONS:
            return QualityReporter._display_locations(locations)

        by_sheet = OrderedDict()
        for location in locations:
            by_sheet.setdefault(location.sheet, []).append(location)

        total = total_count or len(locations)

        def range_label(sheet_locations: List[CellLocation]) -> str:
            rows = [location.row for location in sheet_locations]
            cols = [location.col for location in sheet_locations]
            start = f"{get_column_letter(min(cols))}{min(rows)}"
            end = f"{get_column_letter(max(cols))}{max(rows)}"
            return start if start == end else f"{start}:{end}"

        if len(by_sheet) == 1:
            sheet, sheet_locations = next(iter(by_sheet.items()))
            return sheet, f"{range_label(sheet_locations)}（共核对{total}个位置）"

        summary = "；".join(
            f"{sheet}!{range_label(sheet_locations)}"
            for sheet, sheet_locations in by_sheet.items()
        )
        return "多个Sheet", f"{summary}（共核对{total}个位置）"

    @staticmethod
    def _first_hyperlink_target(sheet: str, address: str) -> tuple:
        """合并地址展示时，超链接跳转到第一个问题单元格。"""
        if sheet == "多个Sheet" and "!" in address:
            first = address.split("；", 1)[0]
            return tuple(first.rsplit("!", 1))
        first_address = re.split(r"[、；]", address, maxsplit=1)[0]
        return sheet, first_address

    @staticmethod
    def _logic_summary_from_detail(detail_text: str) -> dict:
        summary = {"data": "", "method": "", "conclusion": ""}
        for line in str(detail_text or "").splitlines():
            line = line.strip()
            if line.startswith("取数："):
                summary["data"] = line.replace("取数：", "", 1)
            elif line.startswith("方法："):
                summary["method"] = line.replace("方法：", "", 1)
            elif line.startswith("结论："):
                summary["conclusion"] = line.replace("结论：", "", 1)
        return summary

    @staticmethod
    def _report_data_text(summary_data: str, sheet: str, address: str, actual: str = "") -> str:
        if sheet and address:
            return QualityReporter._combined_location(sheet, address)
        return summary_data

    @staticmethod
    def _combined_location(sheet: str, address: str) -> str:
        if sheet == "多个Sheet" and "!" in address:
            return address
        return f"{sheet}!{address}"

    def _add_check_logic_sheet(self, wb, results: List[CheckResult],
                                execution_records: Optional[List[RuleExecutionRecord]] = None):
        """添加核对逻辑Sheet，展示每条规则的执行详情"""
        sheet_name = "核对逻辑"
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

        ws = wb.create_sheet(sheet_name, 1)

        # 标题
        ws['A1'] = "质检核对逻辑详情"
        ws['A1'].font = Font(size=16, bold=True)
        ws['A2'] = ""

        # 表头
        headers = [
            "序号", "规则ID", "规则名称", "执行状态", "取数", "核对方法", "结论", "定位"
        ]
        header_row = 4
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col, value=header)
            cell.font = Font(bold=True)

        current_row = header_row + 1
        row_index = 1
        seen_rules = set()

        # 收集所有规则及其详情
        rule_details = []

        # 1. 先处理有结果的规则（FAILED / 问题项）
        for result in results:
            rule_id = result.rule_id
            seen_rules.add(rule_id)
            sheet, address = self._display_location(result)
            summary = self._logic_summary_from_detail(result.execution_detail)
            if not any(summary.values()):
                summary = build_execution_summary(
                    rule_id,
                    message=result.message,
                    expected=result.expected or "",
                    actual=result.actual or "",
                    sheet=sheet,
                    cell=address,
                    status="FAILED",
                ) or {}

            data_text = self._report_data_text(
                summary.get("data", ""),
                sheet,
                address,
                result.actual or "",
            )

            rule_details.append({
                "rule_id": rule_id,
                "rule_name": result.rule_name,
                "status": "发现问题",
                "data": data_text,
                "method": summary.get("method", ""),
                "conclusion": summary.get("conclusion", result.message or "不通过"),
                "location": self._combined_location(sheet, address) if sheet and address else "",
                "hyperlink_sheet": sheet,
                "hyperlink_address": address,
            })

        # 2. 再处理执行记录中未出现过的规则（PASSED / SKIPPED / AI_DISABLED）
        if execution_records:
            for er in execution_records:
                if er.rule_id in seen_rules:
                    # 已有详细结果，跳过（但如果是PASSED且有独立execution_detail也添加）
                    if er.status == RuleStatus.PASSED and er.issues:
                        continue
                    # 如果同规则有FAILED结果，PASSED的不再重复
                    continue
                seen_rules.add(er.rule_id)
                execution_detail = er.execution_detail or ""
                summary = self._logic_summary_from_detail(execution_detail)
                if not any(summary.values()):
                    summary = build_execution_summary(
                        er.rule_id,
                        message=er.skipped_reason or "",
                        status=er.status.name,
                    ) or {}

                # 状态标签
                if er.status == RuleStatus.PASSED:
                    status_label = "检查通过"
                elif er.status == RuleStatus.SKIPPED:
                    status_label = "未检查"
                    summary["conclusion"] = er.skipped_reason or summary.get("conclusion", "未检查")
                elif er.status == RuleStatus.AI_DISABLED:
                    status_label = "未调用AI"
                    summary["conclusion"] = er.skipped_reason or summary.get("conclusion", "未调用AI")
                else:
                    status_label = er.status.value

                rule_details.append({
                    "rule_id": er.rule_id,
                    "rule_name": er.rule_name,
                    "status": status_label,
                    "data": summary.get("data", ""),
                    "method": summary.get("method", ""),
                    "conclusion": summary.get("conclusion", ""),
                    "location": self._logic_record_location(er),
                    "hyperlink_sheet": self._display_record_evidence(er)[2],
                    "hyperlink_address": self._display_record_evidence(er)[3],
                })

        # 3. 写入数据
        for detail in rule_details:
            ws.cell(row=current_row, column=1, value=row_index)
            ws.cell(row=current_row, column=2, value=detail["rule_id"])
            ws.cell(row=current_row, column=3, value=detail["rule_name"])
            ws.cell(row=current_row, column=4, value=detail["status"])
            ws.cell(row=current_row, column=5, value=detail["data"])
            ws.cell(row=current_row, column=6, value=detail["method"])
            ws.cell(row=current_row, column=7, value=detail["conclusion"])
            addr_cell = ws.cell(row=current_row, column=8, value=detail["location"])
            if detail["hyperlink_sheet"] and detail["hyperlink_address"]:
                try:
                    hyperlink_sheet, hyperlink_address = self._first_hyperlink_target(
                        detail["hyperlink_sheet"], detail["hyperlink_address"]
                    )
                    safe_sheet = hyperlink_sheet.replace("'", "''")
                    addr_cell.hyperlink = f"#'{safe_sheet}'!{hyperlink_address}"
                    addr_cell.font = self.LINK_FONT
                except Exception:
                    pass
            current_row += 1
            row_index += 1

        # 调整列宽
        column_widths = [6, 12, 25, 14, 45, 42, 50, 28]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        self._style_logic_sheet(ws, header_row, current_row - 1)

    def _logic_record_location(self, record: RuleExecutionRecord) -> str:
        sheet, address, _, _ = self._display_record_evidence(record)
        if sheet and address:
            return self._combined_location(sheet, address)
        return address or sheet
