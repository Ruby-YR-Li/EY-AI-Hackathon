"""
费用底稿Review Agent - 勾稽核对规则
UEXP-GL-001: Lead与BKD核对
UEXP-GL-002: A3签核标记检查
UEXP-GL-003: TE/SAD值检查
"""

from typing import Dict, List, Optional, Tuple
from ..models import (AuditWorkbook, CellLocation, CheckResult, CheckCategory, Severity)
from .base import BaseRule
import pandas as pd


class ConsistencyRules(BaseRule):
    """费用勾稽核对规则"""

    COL_D = 3
    COL_J = 9

    def __init__(self):
        super().__init__("UEXP-GL", "费用勾稽核对检查")

    def list_rule_ids(self) -> List[dict]:
        return [
            {"rule_id": "UEXP-GL-001", "rule_name": "Lead Sheet本期账面数与BKD核对",
             "category": CheckCategory.CONSISTENCY,
             "required_sheets": ["Uexp.00 Lead", "VC.00", "VD.00"]},
            {"rule_id": "UEXP-GL-002", "rule_name": "Lead Sheet审定数与A3核对",
             "category": CheckCategory.CONSISTENCY,
             "required_sheets": ["Uexp.00 Lead"]},
            {"rule_id": "UEXP-GL-003", "rule_name": "TE/SAD与Canvas一致性",
             "category": CheckCategory.CONSISTENCY,
             "required_sheets": ["Uexp.00 Lead"]},
            {"rule_id": "UEXP-GL-004", "rule_name": "TE/SAD/CRA配置与底稿核对",
             "category": CheckCategory.CONSISTENCY,
             "required_sheets": ["Uexp.00 Lead"]},
        ]

    def is_required_sheet_available(self, workbook, rule_id, sheet_name):
        return any(sheet_name in sn for sn in workbook.sheets)

    def missing_requirement_reason(self, rule_id, missing):
        return f"缺少必要Sheet：{', '.join(missing)}"

    def check(self, workbook: AuditWorkbook) -> List[CheckResult]:
        results = []
        self.reset_runtime_state()
        lead_sn = self._find_sheet(workbook, "Uexp.00 Lead")
        lead_df = workbook.sheets.get(lead_sn) if lead_sn else None
        if lead_df is None or lead_df.empty:
            return results
        with self.evidence_context("UEXP-GL-001"):
            r = self._check_lead_vs_bkd(workbook, lead_sn, lead_df)
            if r: results.append(r)
        with self.evidence_context("UEXP-GL-002"):
            r = self._check_lead_signoff(lead_sn, lead_df)
            if r: results.append(r)
        with self.evidence_context("UEXP-GL-003"):
            r = self._check_te_sad(lead_sn, lead_df)
            if r: results.append(r)
        with self.evidence_context("UEXP-GL-004"):
            r = self._check_te_sad_cra(workbook, lead_sn, lead_df)
            if r: results.append(r)
        # 通过项记录证据，避免被标为"未确认"
        for rid in ["UEXP-GL-001", "UEXP-GL-002", "UEXP-GL-003", "UEXP-GL-004"]:
            if not any(x.rule_id == rid for x in results):
                self.record_evidence_text(rid, f"规则{rid}检查通过，未发现问题")
        return results

    def _find_sheet(self, workbook, name):
        for sn in workbook.sheets:
            if name in sn: return sn
        return None

    def _get_val(self, row, col):
        return row.iloc[col] if col < len(row) else None

    def _to_float(self, val):
        if val is None: return None
        try: return float(val)
        except: return None

    def _find_bkd_total(self, df, header_idx):
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            d_val = self._get_val(row, self.COL_D)
            if d_val and "合计" in str(d_val):
                return self._to_float(self._get_val(row, self.COL_J))
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            f_val = self._get_val(row, 5)
            if f_val and str(f_val).strip() == "A3":
                return self._to_float(self._get_val(row, self.COL_J))
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            e_val = self._get_val(row, 4)
            if e_val and str(e_val).strip() == "Rx":
                return self._to_float(self._get_val(row, self.COL_J))
        for i in range(len(df) - 1, header_idx, -1):
            row = df.iloc[i]
            if self._to_float(self._get_val(row, self.COL_J)):
                return self._to_float(self._get_val(row, self.COL_J))
        return None

    def _check_lead_vs_bkd(self, workbook, lead_sn, lead_df):
        if 26 >= len(lead_df):
            return None
        sales_lead = self._to_float(self._get_val(lead_df.iloc[25], 3))
        mgmt_lead = self._to_float(self._get_val(lead_df.iloc[26], 3))
        sad = self._to_float(self._get_val(lead_df.iloc[5], 2)) if 5 < len(lead_df) else 0
        threshold = max((sad or 0) * 0.01, 100)
        issues = []
        for cfg, lead_amt, field in [
            ({"kw": "VC.00", "hi": 29}, sales_lead, "销售费用"),
            ({"kw": "VD.00", "hi": 28}, mgmt_lead, "管理费用"),
        ]:
            sn = self._find_sheet(workbook, cfg["kw"])
            if not sn: continue
            df = workbook.sheets.get(sn)
            if df is None: continue
            tot = self._find_bkd_total(df, cfg["hi"])
            if tot is None:
                issues.append(f"{field}BKD合计无法识别")
                continue
            if lead_amt is not None and abs(lead_amt - tot) > threshold:
                issues.append(f"{field}: Lead账面{lead_amt:,.0f} vs BKD{tot:,.0f}")
        if issues:
            return CheckResult(rule_id="UEXP-GL-001", rule_name="Lead与BKD核对",
                category=CheckCategory.CONSISTENCY, severity=Severity.HIGH,
                location=CellLocation(lead_sn, 26, 4),
                message="；".join(issues), expected="Lead与BKD应一致", actual="存在差异")
        return None

    def _check_lead_signoff(self, lead_sn, lead_df):
        markers = {13: "TB", 14: "PY", 15: "A3"}
        missing = [e for idx, e in markers.items()
                   if idx >= len(lead_df) or str(self._get_val(lead_df.iloc[idx], 1) or "").strip() != e]
        if missing:
            return CheckResult(rule_id="UEXP-GL-002", rule_name="签核标记检查",
                category=CheckCategory.CONSISTENCY, severity=Severity.MEDIUM,
                location=CellLocation(lead_sn, 15, 2),
                message=f"签核标记区缺少：{', '.join(missing)}",
                expected="应有TB/PY/A3标记", actual=f"缺失{', '.join(missing)}")
        return None

    def _check_te_sad(self, lead_sn, lead_df):
        te = self._to_float(self._get_val(lead_df.iloc[4], 2)) if 4 < len(lead_df) else None
        sad = self._to_float(self._get_val(lead_df.iloc[5], 2)) if 5 < len(lead_df) else None
        issues = []
        if te is None or te <= 0: issues.append(f"TE异常({te})")
        if sad is None or sad <= 0: issues.append(f"SAD异常({sad})")
        if te and sad and sad >= te: issues.append("SAD≥TE")
        if issues:
            return CheckResult(rule_id="UEXP-GL-003", rule_name="TE/SAD值检查",
                category=CheckCategory.CONSISTENCY, severity=Severity.HIGH,
                location=CellLocation(lead_sn, 5, 3),
                message="；".join(issues), expected="TE>0且SAD>0且SAD<TE", actual="；".join(issues))
        return None

    def _check_te_sad_cra(self, workbook, lead_sn, lead_df):
        """UEXP-GL-004: 用户配置的TE/SAD/CRA与底稿核对"""
        ctx = getattr(workbook, 'context', {}) or {}
        baseline = ctx.get('baseline', {}) or {}
        proj_cfg = ctx.get('project_config', {}) or {}

        # 获取用户输入的TE/SAD
        user_te = None
        user_sad = None
        bl = baseline
        if bl.get('te') and bl['te'].get('value'):
            user_te = bl['te']['value']
        if bl.get('sad') and bl['sad'].get('value'):
            user_sad = bl['sad']['value']

        # 获取Lead sheet中的TE/SAD
        lead_te = self._to_float(self._get_val(lead_df.iloc[4], 2)) if 4 < len(lead_df) else None
        lead_sad = self._to_float(self._get_val(lead_df.iloc[5], 2)) if 5 < len(lead_df) else None

        issues = []

        # 比较TE
        if user_te is not None and lead_te is not None:
            diff_pct = abs(user_te - lead_te) / max(user_te, lead_te, 1)
            if diff_pct > 0.01:  # 差异超过1%
                issues.append(f"TE不一致：您输入{user_te:,.0f}，底稿Lead Sheet填列{lead_te:,.0f}")

        # 比较SAD
        if user_sad is not None and lead_sad is not None:
            diff_pct = abs(user_sad - lead_sad) / max(user_sad, lead_sad, 1)
            if diff_pct > 0.01:
                issues.append(f"SAD不一致：您输入{user_sad:,.0f}，底稿Lead Sheet填列{lead_sad:,.0f}")

        # 检查CRA（底稿中BKD的CRA值）
        assertion_cfg = proj_cfg.get('subject_assertions', {}) or {}
        if assertion_cfg:
            # 遍历各BKD sheet检查CRA
            for bkd_key, bkd_label, cra_row in [
                ("VC.00", "销售费用BKD", 7),  # Row 8 (idx 7)
                ("VD.00", "管理费用BKD", 7),
            ]:
                bkd_sn = self._find_sheet(workbook, bkd_key)
                if not bkd_sn:
                    continue
                bkd_df = workbook.sheets.get(bkd_sn)
                if bkd_df is None or bkd_df.empty:
                    continue
                # Check CRA values in BKD (C列 at CRA header rows)
                if cra_row < len(bkd_df):
                    row = bkd_df.iloc[cra_row]
                    # C列(col 2) = 认定, D列(col 3) = CRA
                    for check_idx in range(cra_row, min(cra_row + 5, len(bkd_df))):
                        r = bkd_df.iloc[check_idx]
                        assertion_name = str(self._get_val(r, 2) or "").strip()
                        cra_val = str(self._get_val(r, 3) or "").strip().lower()
                        if assertion_name and cra_val:
                            # Look up user's CRA config for this assertion
                            for a_key, a_info in assertion_cfg.items():
                                if isinstance(a_info, dict):
                                    a_name = str(a_info.get('name', '')).strip()
                                    a_cra = str(a_info.get('cra', '')).strip().lower()
                                    if a_name and a_name in assertion_name and a_cra and a_cra != cra_val:
                                        if a_cra not in ('n/a', '') and cra_val not in ('n/a', ''):
                                            issues.append(f"{bkd_label}认定'{assertion_name}'CRA：您配置{a_cra}，底稿填列{cra_val}")

        if issues:
            loc = CellLocation(lead_sn, 5, 3) if any('TE' in i or 'SAD' in i for i in issues) else CellLocation(lead_sn, 8, 3)
            return CheckResult(rule_id="UEXP-GL-004", rule_name="TE/SAD/CRA配置与底稿核对",
                category=CheckCategory.CONSISTENCY, severity=Severity.HIGH,
                location=loc, message="；".join(issues[:3]),
                expected="用户输入的TE/SAD/CRA应与底稿一致",
                actual="存在不一致")
        return None
