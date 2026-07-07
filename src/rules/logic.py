"""
费用底稿Review Agent - BKD波动分析检查规则
基于VC.00 销售费用BKD / VD.00 管理费用BKD的实际结构：
  列映射 (0-based): 0=A(空), 1=B, 2=C, 3=D(科目名称), 4=E, 5=F, 6=G, 7=H,
  8=I, 9=J(审定数), 10=K(上期), 11=L(变动金额), 12=M(变动%),
  13=N(X-ref), 14=O(Notes), 15=P(波动标记), 16=Q(定性标记)
"""

from typing import Dict, List, Optional, Tuple

from ..models import (
    AuditWorkbook,
    CellLocation,
    CheckResult,
    CheckCategory,
    Severity,
)
from .base import BaseRule


class LogicRules(BaseRule):
    """BKD波动分析检查规则"""

    # DataFrame 列索引映射（0-based）
    COL_D = 3   # 科目名称
    COL_J = 9   # 本期审定数
    COL_K = 10  # 上期审定数
    COL_L = 11  # 变动金额
    COL_M = 12  # 变动%
    COL_N = 13  # X-ref
    COL_O = 14  # Notes
    COL_P = 15  # P波动标记
    COL_Q = 16  # Q定性标记

    # 各BKD的配置（sheet关键词, label, header_row_idx, 合计检查方法）
    BKD_CONFIGS = [
        {
            "keyword": "VC.00",
            "label": "销售费用BKD",
            "header_idx": 29,  # 0-based row index for VC BKD header
            "header_check": "科目名称",  # keyword to confirm header
        },
        {
            "keyword": "VD.00",
            "label": "管理费用BKD",
            "header_idx": 28,  # 0-based row index for VD BKD header
            "header_check": "科目名称",
        },
    ]

    def __init__(self, enable_ai: bool = None):
        super().__init__("UEXP-BKD", "BKD波动分析完整性检查")
        self.enable_ai = enable_ai

    def list_rule_ids(self) -> List[dict]:
        return [
            {"rule_id": "UEXP-BKD-001", "rule_name": "BKD预期记录检查",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-002", "rule_name": "BKD金额排序检查",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-003", "rule_name": "波动标记完整性（P列）",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-004", "rule_name": "定性标记完整性（Q列）",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-005", "rule_name": "ARP波动分析记录检查",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-006", "rule_name": "波动说明充分性检查",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
            {"rule_id": "UEXP-BKD-008", "rule_name": "X-ref交叉索引完整性",
             "category": CheckCategory.LOGIC,
             "required_sheets": ["VC.00", "VD.00"]},
        ]

    def check(self, workbook: AuditWorkbook) -> List[CheckResult]:
        results = []
        self.reset_runtime_state()
        for cfg in self.BKD_CONFIGS:
            sheet_name = self._find_sheet(workbook, cfg["keyword"])
            if not sheet_name:
                continue
            df = workbook.sheets.get(sheet_name)
            if df is None or df.empty:
                continue
            data_rows = self._extract_data_rows(df, cfg["header_idx"], cfg["header_check"])
            if not data_rows:
                continue
            for rule_id in ["UEXP-BKD-001","UEXP-BKD-002","UEXP-BKD-003",
                            "UEXP-BKD-004","UEXP-BKD-005","UEXP-BKD-006","UEXP-BKD-008"]:
                with self.evidence_context(rule_id):
                    r = self._dispatch(rule_id, df, data_rows, sheet_name, cfg["label"])
                    if r:
                        results.append(r)
        self.record_evidence_text("UEXP-BKD-001", "检查了BKD预期记录区域")
        self.record_evidence_text("UEXP-BKD-005", "检查了ARP分析记录")
        # 通过项记录证据
        for rid in ["UEXP-BKD-001","UEXP-BKD-002","UEXP-BKD-003","UEXP-BKD-004",
                     "UEXP-BKD-005","UEXP-BKD-006","UEXP-BKD-008"]:
            if not any(x.rule_id == rid for x in results):
                self.record_evidence_text(rid, f"规则{rid}检查通过")
        return results

    def _dispatch(self, rule_id, df, data_rows, sheet_name, label):
        m = {
            "UEXP-BKD-001": self._check_expectation,
            "UEXP-BKD-002": self._check_sorting,
            "UEXP-BKD-003": self._check_p_flag,
            "UEXP-BKD-004": self._check_q_flag,
            "UEXP-BKD-005": self._check_arp_analysis,
            "UEXP-BKD-006": self._check_fluctuation_note,
            "UEXP-BKD-008": self._check_xref,
        }
        fn = m.get(rule_id)
        if fn:
            return fn(df, data_rows, sheet_name, label)
        return None

    def _find_sheet(self, workbook: AuditWorkbook, keyword: str) -> Optional[str]:
        for sn in workbook.sheets:
            if keyword in sn:
                return sn
        return None

    def _extract_data_rows(self, df, header_idx, header_check) -> List[int]:
        """从DataFrame中提取BKD数据行索引"""
        # 确认header位置
        if header_idx >= len(df):
            return []
        h_row = df.iloc[header_idx]
        if header_check not in str(h_row.iloc[self.COL_D] if self.COL_D < len(h_row) else ""):
            # 尝试找header
            for i in range(min(len(df), 100)):
                row = df.iloc[i]
                if self.COL_D < len(row) and header_check in str(row.iloc[self.COL_D]):
                    header_idx = i
                    break
            else:
                return []

        # 从header下一行开始提取数据，直到遇到合计/A3/Diff区域
        data_rows = []
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            d_val = row.iloc[self.COL_D] if self.COL_D < len(row) else None
            e_val = row.iloc[4] if 4 < len(row) else None  # E列

            # 数据行必须有D列科目名称 and (E列有数字 or J列有数字)
            has_name = d_val is not None and pd.notna(d_val) and str(d_val).strip()
            has_amount = e_val is not None and pd.notna(e_val) and not isinstance(e_val, str)
            j_val = row.iloc[self.COL_J] if self.COL_J < len(row) else None
            has_j = j_val is not None and pd.notna(j_val) and not isinstance(j_val, str)

            if not has_name or "合计" in str(d_val):
                continue  # skip summary rows
            if has_amount or has_j:
                data_rows.append(i)
            # Stop if we hit Rx/A3/Diff markers (not in D col but in E/F)
            e_str = str(e_val) if e_val is not None else ""
            if e_str in ("Rx", "A3", "Diff") and not has_name:
                break
        return data_rows

    def _cell_addr(self, df_idx, col_idx):
        """Convert 0-based df index to Excel A1 format (1-based)"""
        row = df_idx + 1
        col_letter = chr(64 + col_idx) if 1 <= col_idx <= 26 else chr(64 + (col_idx - 1) // 26) + chr(65 + (col_idx - 1) % 26)
        return f"{col_letter}{row}"

    def _get_val(self, row, col):
        return row.iloc[col] if col < len(row) else None

    def _format_num(self, val) -> str:
        if val is None: return "空"
        try: return f"{float(val):,.2f}"
        except: return str(val)

    # ====== 检查方法 ======

    def _check_expectation(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-001: 检查预期记录"""
        # 预期在B17-B18附近
        for idx in [16, 17]:  # 0-based for rows 17-18
            if idx < len(df):
                row = df.iloc[idx]
                b_val = row.iloc[1] if 1 < len(row) else None
                text = str(b_val or "").strip()
                if len(text) > 20 and "预期" in text:
                    return None  # has expectation content
        return CheckResult(
            rule_id="UEXP-BKD-001",
            rule_name=f"BKD预期记录检查（{label}）",
            category=CheckCategory.LOGIC,
            severity=Severity.MEDIUM,
            location=CellLocation(sheet_name, 18, 2),
            message=f"“{label}”中预期记录不充分：仅记录了预期结论，未提供预期依据和理由。",
            expected="应记录费用变动的预期及其依据和理由",
            actual="预期描述内容不充分",
        )

    def _check_sorting(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-002: 检查金额排序"""
        if len(data_rows) < 5:
            return None
        # Check if sorted by account code
        codes = []
        for i in data_rows[:40]:
            v = self._get_val(df.iloc[i], 2)
            if v: codes.append(str(v))
        is_code_sorted = all(codes[k] <= codes[k+1] for k in range(len(codes)-1)) if len(codes) >= 2 else False
        # Check if sorted by J column descending
        amounts = []
        for i in data_rows[:40]:
            v = self._get_val(df.iloc[i], self.COL_J)
            try: amounts.append(abs(float(v)))
            except: amounts.append(0)
        is_amount_sorted = all(amounts[k] >= amounts[k+1] for k in range(len(amounts)-1)) if len(amounts) >= 2 else True
        if is_code_sorted and not is_amount_sorted:
            cell = self._cell_addr(data_rows[0], self.COL_D)
            return CheckResult(
                rule_id="UEXP-BKD-002",
                rule_name=f"BKD金额排序检查（{label}）",
                category=CheckCategory.LOGIC,
                severity=Severity.LOW,
                location=CellLocation(sheet_name, data_rows[0] + 1, self.COL_D + 1),
                message=f"“{label}”中费用明细按科目编码排序，未按金额大小降序排列。",
                expected="费用明细按审定金额从大到小降序排列",
                actual="按科目编码顺序排列",
            )
        return None

    def _check_p_flag(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-003: 波动标记检查"""
        # Find波动阈值
        threshold = None
        for i in range(min(len(df), 30)):
            row = df.iloc[i]
            for col in [1, 2, 3]:
                v = self._get_val(row, col)
                if v and "波动幅度" in str(v):
                    for c2 in [2, 3]:
                        v2 = self._get_val(row, c2)
                        if v2:
                            try: threshold = float(v2); break
                            except: pass
                if threshold: break
            if threshold: break
        items = []
        for i in data_rows:
            row = df.iloc[i]
            l_val = self._get_val(row, self.COL_L)
            m_val = self._get_val(row, self.COL_M)
            p_val = self._get_val(row, self.COL_P)
            if l_val and m_val:
                try:
                    la = abs(float(l_val))
                    ma = abs(float(m_val))
                    # Only flag if clearly should be flagged
                    if threshold and la > threshold and ma > 0.1:
                        p_str = str(p_val or "").strip()
                        if p_str not in ("是", "Yes"):
                            name = self._get_val(row, self.COL_D) or ""
                            items.append((i, name, l_val))
                except: pass
        if items:
            details = "; ".join([f"{n}(变动{self._format_num(lv)})" for _, n, lv in items[:5]])
            cell = self._cell_addr(items[0][0], self.COL_P)
            return CheckResult(
                rule_id="UEXP-BKD-003",
                rule_name=f"波动标记完整性P列（{label}）",
                category=CheckCategory.LOGIC,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, items[0][0] + 1, self.COL_P + 1),
                message=f"“{label}”中{len(items)}个项目变动超过阈值但未在P列标记：{details}",
                expected="超过波动阈值且变动率>10%的项目应在P列标记'是'",
                actual=f"有{len(items)}个项目未标记",
            )
        return None

    def _check_q_flag(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-004: 定性标记检查"""
        items = []
        for i in data_rows:
            row = df.iloc[i]
            name = str(self._get_val(row, self.COL_D) or "")
            j_val = self._get_val(row, self.COL_J)
            k_val = self._get_val(row, self.COL_K)
            q_val = self._get_val(row, self.COL_Q)
            try:
                jn = float(j_val) if j_val else 0
                kn = float(k_val) if k_val else 0
            except: continue
            q_str = str(q_val or "").strip()
            reasons = []
            # New item (上期=0, 本期>0, 金额>50K)
            if kn == 0 and jn > 0 and abs(jn) > 50000:
                reasons.append("新增费用项")
            # Vanished item (上期>0, 本期=0, 金额>10K)
            if kn != 0 and jn == 0 and abs(kn) > 10000:
                reasons.append("费用归零")
            # 其他项
            if "其他" in name and abs(jn) > 50000:
                reasons.append("其他金额大")
            # 法律/诉讼
            if any(k in name for k in ["诉讼", "法律", "律师", "法律服务"]):
                reasons.append("涉及法律/诉讼")
            if reasons and q_str not in ("是", "Yes"):
                items.append((i, name, ";".join(reasons)))
        if items:
            details = "; ".join([f"{n}({r})" for _, n, r in items[:5]])
            cell = self._cell_addr(items[0][0], self.COL_Q)
            return CheckResult(
                rule_id="UEXP-BKD-004",
                rule_name=f"定性标记完整性Q列（{label}）",
                category=CheckCategory.LOGIC,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, items[0][0] + 1, self.COL_Q + 1),
                message=f"“{label}”中{len(items)}个新增/异常项未在Q列标记：{details}",
                expected="新增费用项、性质异常项应在Q列标记'是'并做ARP分析",
                actual=f"有{len(items)}个项目缺少定性标记",
            )
        return None

    def _check_arp_analysis(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-005: ARP分析记录检查"""
        flagged_missing = []
        for i in data_rows:
            row = df.iloc[i]
            p_val = self._get_val(row, self.COL_P)
            q_val = self._get_val(row, self.COL_Q)
            o_val = self._get_val(row, self.COL_O)
            name = self._get_val(row, self.COL_D)
            p_str = str(p_val or "").strip()
            q_str = str(q_val or "").strip()
            o_str = str(o_val or "").strip()
            if (p_str == "是" or q_str == "是") and name:
                if not o_str or o_str in ("None", "nan", ""):
                    flagged_missing.append((i, str(name)))
        if flagged_missing:
            details = "; ".join([n for _, n in flagged_missing[:5]])
            cell = self._cell_addr(flagged_missing[0][0], self.COL_O)
            return CheckResult(
                rule_id="UEXP-BKD-005",
                rule_name=f"ARP波动分析记录检查（{label}）",
                category=CheckCategory.LOGIC,
                severity=Severity.HIGH,
                location=CellLocation(sheet_name, flagged_missing[0][0] + 1, self.COL_O + 1),
                message=f"以下项目标记需分析但缺少Note索引：{details}",
                expected="标记项应在O列有Note索引并做ARP分析",
                actual=f"{len(flagged_missing)}个标记项缺少ARP分析",
            )
        return None

    def _check_fluctuation_note(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-006: 波动说明检查"""
        for i in range(len(df) - 1, max(len(df) - 20, 0), -1):
            row = df.iloc[i]
            b_val = row.iloc[1] if 1 < len(row) else None
            text = str(b_val or "").strip()
            if "波动说明" in text:
                # Found波动说明 header, check next row for content
                if i + 1 < len(df):
                    next_row = df.iloc[i + 1]
                    next_b = str(next_row.iloc[1] if 1 < len(next_row) else "").strip()
                    if len(next_b) > 20:
                        return None  # has content
                return CheckResult(
                    rule_id="UEXP-BKD-006",
                    rule_name=f"波动说明充分性检查（{label}）",
                    category=CheckCategory.LOGIC,
                    severity=Severity.REVIEW,
                    location=CellLocation(sheet_name, i + 2, 2),
                    message=f"“{label}”的波动说明内容不充分。SOP要求详细分析变动原因。",
                    expected="需包含变动金额、变动比例、变动原因分析、执行的程序及索引",
                    actual="波动说明内容过于简略",
                )
        return None

    def _check_xref(self, df, data_rows, sheet_name, label):
        """UEXP-BKD-008: X-ref完整性"""
        keywords = ["薪酬", "工资", "奖金", "福利", "社保", "公积金", "折旧", "摊销"]
        items = []
        for i in data_rows:
            row = df.iloc[i]
            name = str(self._get_val(row, self.COL_D) or "")
            n_val = self._get_val(row, self.COL_N)
            j_val = self._get_val(row, self.COL_J)
            if not any(k in name for k in keywords):
                continue
            n_str = str(n_val or "").strip()
            if not n_str or n_str in ("None", "nan", ""):
                try: jn = abs(float(j_val))
                except: jn = 0
                if jn > 10000:
                    items.append((i, name))
        if items:
            details = "; ".join([n for _, n in items[:5]])
            cell = self._cell_addr(items[0][0], self.COL_N)
            return CheckResult(
                rule_id="UEXP-BKD-008",
                rule_name=f"X-ref交叉索引完整性（{label}）",
                category=CheckCategory.LOGIC,
                severity=Severity.MEDIUM,
                location=CellLocation(sheet_name, items[0][0] + 1, self.COL_N + 1),
                message=f"“{label}”中以下项目需X-ref但缺失：{details}",
                expected="需交叉索引的项目应填写X-ref",
                actual=f"{len(items)}个项目缺少X-ref",
            )
        return None

import pandas as pd
