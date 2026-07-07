"""
费用底稿Review Agent - TOD和截止性测试检查规则
涵盖：TOD样本总体SCOT区分、X-ref完整性、剔除项检查、关键项检查
以及截止测试期间/策略/样本合理性检查
"""

from typing import Dict, List, Optional

from ..models import (
    AuditWorkbook,
    CellLocation,
    CheckResult,
    CheckCategory,
    Severity,
)
from .base import BaseRule


class CompletenessRules(BaseRule):
    """TOD和截止性测试检查规则"""

    def __init__(self, enable_ai: bool = None):
        super().__init__("UEXP-TODCO", "TOD与截止性测试完整性检查")
        self.enable_ai = enable_ai

    # ── 规则注册 ──────────────────────────────────────────────

    def list_rule_ids(self) -> List[dict]:
        return [
            {
                "rule_id": "UEXP-TOD-001",
                "rule_name": "样本总体SCOT区分检查",
                "category": CheckCategory.COMPLETENESS,
                "required_sheets": ["VC&VD.01.2 TOD"],
            },
            {
                "rule_id": "UEXP-TOD-002",
                "rule_name": "TOD抽样策略X-ref完整性",
                "category": CheckCategory.COMPLETENESS,
                "required_sheets": ["VC&VD.01.2 TOD"],
            },
            {
                "rule_id": "UEXP-CO-001",
                "rule_name": "截止测试期间确定依据",
                "category": CheckCategory.COMPLETENESS,
                "required_sheets": ["VC&VD.01.4 截止性测试"],
            },
            {
                "rule_id": "UEXP-CO-002",
                "rule_name": "截止测试样本策略记录",
                "category": CheckCategory.COMPLETENESS,
                "required_sheets": ["VC&VD.01.4 截止性测试"],
            },
            {
                "rule_id": "UEXP-CO-003",
                "rule_name": "截止测试样本金额合理性",
                "category": CheckCategory.COMPLETENESS,
                "required_sheets": ["VC&VD.01.4 截止性测试"],
            },
        ]

    def missing_requirement_reason(self, rule_id: str, missing: List[str]) -> str:
        if "TOD" in rule_id:
            return "未找到TOD底稿Sheet，可能未执行详细测试"
        if "截止" in rule_id:
            return "未找到截止性测试Sheet，可能未执行截止测试"
        return "未找到所需Sheet"

    # ── 主入口 ────────────────────────────────────────────────

    def check(self, workbook: AuditWorkbook) -> List[CheckResult]:
        results = []
        self.reset_runtime_state()
        parser = getattr(self, 'parser', None)
        if parser and hasattr(parser, 'evidence_recorder'):
            self.attach_parser(parser)

        # TOD checks
        tod_sheet = self._find_sheet(workbook, "VC&VD.01.2 TOD")
        if tod_sheet:
            df = workbook.sheets.get(tod_sheet)
            if df is not None and not df.empty:
                with self.evidence_context("UEXP-TOD-001"):
                    r = self._check_tod_scot(df, tod_sheet)
                    if r: results.append(r)
                with self.evidence_context("UEXP-TOD-002"):
                    r = self._check_tod_xref(df, tod_sheet)
                    if r: results.append(r)

        # Cutoff checks
        co_sheet = self._find_sheet(workbook, "VC&VD.01.4 截止性测试")
        if co_sheet:
            df = workbook.sheets.get(co_sheet)
            if df is not None and not df.empty:
                with self.evidence_context("UEXP-CO-001"):
                    r = self._check_co_period(df, co_sheet)
                    if r: results.append(r)
                with self.evidence_context("UEXP-CO-002"):
                    r = self._check_co_strategy(df, co_sheet)
                    if r: results.append(r)
                with self.evidence_context("UEXP-CO-003"):
                    r = self._check_co_sample_amount(df, co_sheet)
                    if r: results.append(r)

        self.record_evidence_text("UEXP-TOD-001", "检查了TOD样本总体中的SCOT区分")
        self.record_evidence_text("UEXP-CO-001", "检查了截止性测试相关数据")
        # 通过项记录证据
        for rid in ["UEXP-TOD-001","UEXP-TOD-002","UEXP-CO-001","UEXP-CO-002","UEXP-CO-003"]:
            if not any(x.rule_id == rid for x in results):
                self.record_evidence_text(rid, f"规则{rid}检查通过")
        return results

    # ── 辅助方法 ──────────────────────────────────────────────

    def _find_sheet(self, workbook: AuditWorkbook, name: str) -> Optional[str]:
        for sn in workbook.sheets:
            if name in sn:
                return sn
        return None

    def _get_value(self, row, col_idx):
        try:
            return row.iloc[col_idx]
        except (IndexError, TypeError):
            return None

    def _cell_addr(self, row_idx, col_idx):
        col_letter = chr(64 + col_idx) if col_idx <= 26 else chr(64 + (col_idx - 1) // 26) + chr(65 + (col_idx - 1) % 26)
        return f"{col_letter}{row_idx + 1}"

    # ── TOD 检查 ──────────────────────────────────────────────

    def _check_tod_scot(self, df, sheet_name):
        """UEXP-TOD-001: 检查TOD样本总体是否混合了不同SCOT的科目"""
        # Look for 总体名称/样本总体 rows
        pop_names = []
        found_header = False
        for idx in range(min(len(df), 50)):
            row = df.iloc[idx]
            vals = [str(v) for v in row if v is not None and str(v).strip()]
            row_text = " ".join(vals)
            if "总体名称" in row_text or "样本总体" in row_text:
                found_header = True
                continue
            if found_header:
                d_val = self._get_value(row, 3) if len(row) > 3 else None  # D列
                f_val = self._get_value(row, 5) if len(row) > 5 else None  # F列
                if d_val and str(d_val).strip() and f_val is not None:
                    pop_names.append(str(d_val).strip())

        # Check if manufacturing expenses or R&D expenses are included
        invalid_pops = []
        for p in pop_names:
            if any(kw in p for kw in ["制造费用", "研发支出", "研发费用", "营业外"]):
                invalid_pops.append(p)

        if invalid_pops:
            return CheckResult(
                rule_id="UEXP-TOD-001",
                rule_name="样本总体SCOT区分检查",
                category=CheckCategory.COMPLETENESS,
                severity=Severity.HIGH,
                location=CellLocation(sheet_name, 11, 4),
                message=f"TOD样本总体中包含了以下非同一SCOT的科目：{'; '.join(invalid_pops)}。"
                        f"SOP要求：不同交易类别（如制造费用、研发支出等）应区分不同样本总体。",
                expected="样本总体应仅包含相同SCOT的费用类型",
                actual=f"样本总体中包含了{'; '.join(invalid_pops)}",
            )

        return None

    def _check_tod_xref(self, df, sheet_name):
        """UEXP-TOD-002: 检查TOD抽样策略的X-ref完整性"""
        # Check 数据来源 X-ref column (E column)
        missing_xref = []
        found_header = False
        for idx in range(min(len(df), 50)):
            row = df.iloc[idx]
            vals = [str(v) for v in row if v is not None and str(v).strip()]
            row_text = " ".join(vals)
            if "数据来源" in row_text and "X-ref" in row_text:
                found_header = True
                continue
            if found_header:
                d_val = self._get_value(row, 3) if len(row) > 3 else None  # D列
                e_val = self._get_value(row, 4) if len(row) > 4 else None  # E列 - X-ref
                f_val = self._get_value(row, 5) if len(row) > 5 else None  # F列
                if d_val and str(d_val).strip() and f_val is not None:
                    e_str = str(e_val or "").strip()
                    if not e_str or e_str in ("None", "nan", ""):
                        missing_xref.append(str(d_val).strip())

        if missing_xref:
            return CheckResult(
                rule_id="UEXP-TOD-002",
                rule_name="TOD抽样策略X-ref完整性",
                category=CheckCategory.COMPLETENESS,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, 11, 5),
                message=f"TOD抽样策略中以下总体缺少数据来源X-ref：{'; '.join(missing_xref)}",
                expected="每个样本总体应填写数据来源X-ref，如<VC.00销售费用BKD>",
                actual=f"{len(missing_xref)}个总体缺少X-ref",
            )

        return None

    # ── 截止性测试 检查 ─────────────────────────────────────────

    def _check_co_period(self, df, sheet_name):
        """UEXP-CO-001: 检查截止测试期间确定依据"""
        found_reason = False
        for idx in range(min(len(df), 20)):
            row = df.iloc[idx]
            vals = [str(v) for v in row if v is not None and str(v).strip()]
            row_text = " ".join(vals)
            if "测试期间" in row_text or "确定依据" in row_text or "截止" in row_text:
                for ci in range(idx, min(idx + 5, len(df))):
                    cr = df.iloc[ci]
                    cv = [str(v) for v in cr if v not in (None, "") and str(v).strip()]
                    ct = " ".join(str(v) for v in cv)
                    if len(ct) > 30:
                        found_reason = True
                        break

        if not found_reason:
            return CheckResult(
                rule_id="UEXP-CO-001",
                rule_name="截止测试期间确定依据",
                category=CheckCategory.COMPLETENESS,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, 9, 3),
                message="截止性测试未记录期间确定依据或依据不充分。"
                        "SOP要求：应根据被审计单位费用从发生到入账的周期确定截止性测试的期间。",
                expected="应记录确定截止性测试涵盖期间的依据和原因",
                actual="未找到期间确定依据记录",
            )

        return None

    def _check_co_strategy(self, df, sheet_name):
        """UEXP-CO-002: 检查截止测试策略记录"""
        found_strategy = False
        for idx in range(min(len(df), 30)):
            row = df.iloc[idx]
            vals = [str(v) for v in row if v is not None and str(v).strip()]
            row_text = " ".join(vals)
            if "策略" in row_text or "key item" in row_text.lower() or "关键项" in row_text:
                found_strategy = True
                break

        if not found_strategy:
            return CheckResult(
                rule_id="UEXP-CO-002",
                rule_name="截止测试样本策略记录",
                category=CheckCategory.COMPLETENESS,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, 12, 2),
                message="截止性测试未记录选样策略。"
                        "SOP要求：应记录截止性测试样本的选择方法和关键项的确定标准。",
                expected="应记录截止测试样本策略，包括关键项的定量/定性标准",
                actual="未找到选样策略记录",
            )

        return None

    def _check_co_sample_amount(self, df, sheet_name):
        """UEXP-CO-003: 检查截止测试样本金额合理性"""
        # Find sample entries and check amounts
        sample_amounts_before = []
        sample_amounts_after = []
        in_before = False
        in_after = False

        for idx in range(len(df)):
            row = df.iloc[idx]
            vals = [str(v) for v in row if v is not None and str(v).strip()]
            row_text = " ".join(vals)

            if "本期资产负债表日前" in row_text:
                in_before = True
                in_after = False
                continue
            if "本期资产负债表日后" in row_text:
                in_after = True
                in_before = False
                continue

            if (in_before or in_after) and len(row) > 3:
                d_val = self._get_value(row, 3)  # D column - amount
                e_val = self._get_value(row, 4)  # E column - description
                if d_val and e_val:
                    try:
                        amt = float(d_val)
                        if in_before:
                            sample_amounts_before.append(amt)
                        else:
                            sample_amounts_after.append(amt)
                    except (ValueError, TypeError):
                        pass

        if sample_amounts_after:
            max_after = max(sample_amounts_after)
            if max_after < 100000:  # if max sample amount is too small
                return CheckResult(
                    rule_id="UEXP-CO-003",
                    rule_name="截止测试样本金额合理性",
                    category=CheckCategory.COMPLETENESS,
                    severity=Severity.REVIEW,
                    location=CellLocation(sheet_name, 32, 4),
                    message=f"截止性测试资产负债表日后选取的样本金额均较小（最大{max_after:,.2f}）。"
                            f"SOP要求：应首先抽取发生额大于测试阈值的交易作为关键项。",
                    expected="截止测试应选取金额较大的交易作为关键样本",
                    actual=f"资产负债表日后样本最大金额仅{max_after:,.2f}",
                )

        return None
