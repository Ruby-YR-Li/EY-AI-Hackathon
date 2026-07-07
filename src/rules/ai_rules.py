"""
费用底稿Review Agent - AI增强判断规则
涵盖：异常波动识别、费用分类异常、PSP拒绝理由判断等
"""

from typing import Dict, List, Optional

from ..models import (
    AuditWorkbook,
    CellLocation,
    CheckResult,
    CheckCategory,
    Severity,
)
from ..ai_prompts import AI_CHECK_TARGETS
from ..config import ENABLE_AI_CHECK
from .base import BaseRule


class AIRules(BaseRule):
    """AI增强检查规则"""

    def __init__(self, enable_ai: bool = True):
        super().__init__("UEXP-AI", "费用AI增强检查规则")
        self.enable_ai = enable_ai
        self.ai_client = None
        if enable_ai:
            self._init_ai_client()

    def is_enabled(self) -> bool:
        return self.enable_ai and self.ai_client is not None

    def needs_ai(self, rule_id: str) -> bool:
        """表明哪些规则需要AI才能执行"""
        ai_rules = {"UEXP-PSP-002", "UEXP-AI-001", "UEXP-AI-002", "UEXP-AI-003"}
        return rule_id in ai_rules

    def _init_ai_client(self):
        try:
            from ..ai_client import create_ai_client
            self.ai_client = create_ai_client()
        except Exception as e:
            print(f"AI客户端初始化失败: {e}")
            self.enable_ai = False

    # ── 规则注册 ──────────────────────────────────────────────

    def list_rule_ids(self) -> List[dict]:
        rules = [
            {
                "rule_id": "UEXP-PSP-002",
                "rule_name": "PSP拒绝执行理由充分性",
                "category": CheckCategory.AI_REVIEW,
                "required_sheets": ["汇总"],
                "ai_required": True,
            },
            {
                "rule_id": "UEXP-AI-001",
                "rule_name": "BKD异常波动AI辅助判断",
                "category": CheckCategory.AI_REVIEW,
                "required_sheets": ["VC.00 销售费用BKD", "VD.00 管理费用BKD"],
                "ai_required": True,
            },
            {
                "rule_id": "UEXP-AI-002",
                "rule_name": "费用分类异常AI识别",
                "category": CheckCategory.AI_REVIEW,
                "required_sheets": ["VC.00 销售费用BKD", "VD.00 管理费用BKD"],
                "ai_required": True,
            },
            {
                "rule_id": "UEXP-AI-003",
                "rule_name": "法律费用关注性判断",
                "category": CheckCategory.AI_REVIEW,
                "required_sheets": ["VD.00 管理费用BKD"],
                "ai_required": True,
            },
        ]
        return rules

    def list_ai_rule_info(self) -> List[dict]:
        return [
            {"rule_id": "UEXP-PSP-002", "name": "PSP拒绝执行理由充分性", "sheets": ["汇总"]},
            {"rule_id": "UEXP-AI-001", "name": "BKD异常波动AI辅助判断", "sheets": ["VC.00 销售费用BKD"]},
            {"rule_id": "UEXP-AI-002", "name": "费用分类异常AI识别", "sheets": ["VC.00 销售费用BKD"]},
            {"rule_id": "UEXP-AI-003", "name": "法律费用关注性判断", "sheets": ["VD.00 管理费用BKD"]},
        ]

    # ── 主入口 ────────────────────────────────────────────────

    def check(self, workbook: AuditWorkbook) -> List[CheckResult]:
        """执行所有AI规则检查"""
        results = []
        self.reset_runtime_state()
        parser = getattr(self, 'parser', None)
        if parser and hasattr(parser, 'evidence_recorder'):
            self.attach_parser(parser)

        if not self.enable_ai or not self.ai_client:
            for rid in ["UEXP-PSP-002", "UEXP-AI-001", "UEXP-AI-002", "UEXP-AI-003"]:
                self.record_evidence_text(rid, "AI未启用或API密钥未配置")
            return results

        with self.evidence_context("UEXP-PSP-002"):
            r = self._check_psp_rejection("UEXP-PSP-002", workbook)
            if r: results.append(r)

        with self.evidence_context("UEXP-AI-001"):
            r = self._check_fluctuation_ai("UEXP-AI-001", workbook)
            if r: results.append(r)

        with self.evidence_context("UEXP-AI-002"):
            r = self._check_classification_ai("UEXP-AI-002", workbook)
            if r: results.append(r)

        with self.evidence_context("UEXP-AI-003"):
            r = self._check_legal_fees_ai("UEXP-AI-003", workbook)
            if r: results.append(r)

        # 为所有AI规则记录evidence（即使没有发现问题）
        all_ai_rule_ids = ["UEXP-PSP-002", "UEXP-AI-001", "UEXP-AI-002", "UEXP-AI-003"]
        for rid in all_ai_rule_ids:
            has_issue = any(x.rule_id == rid for x in results)
            if has_issue:
                continue  # 有问题时已有evidence
            if not self.enable_ai or not self.ai_client:
                self.record_evidence_text(rid, "AI未启用，跳过检查")
            else:
                self.record_evidence_text(rid, f"规则{rid}已执行AI判断，未发现需关注问题")

        return results

    # ── AI辅助检查方法 ────────────────────────────────────────

    def _ai_judge(self, prompt: str) -> dict:
        """调用AI进行判断"""
        if not self.ai_client:
            return {"has_issue": False, "reason": "AI未配置", "suggestion": ""}
        try:
            result = self.ai_client.judge(prompt)
            if result and not result.is_error:
                return {
                    "has_issue": result.has_issue,
                    "reason": result.reason or "",
                    "suggestion": result.suggestion or "",
                }
        except Exception as e:
            pass
        return {"has_issue": False, "reason": "AI调用失败", "suggestion": ""}

    def _check_psp_rejection(self, rule_id: str, workbook: AuditWorkbook) -> Optional[CheckResult]:
        """UEXP-PSP-002: AI判断PSP拒绝理由充分性"""
        summary_sheet = None
        for sn in workbook.sheets:
            if "汇总" in sn:
                summary_sheet = sn
                break

        if not summary_sheet:
            return None

        df = workbook.sheets.get(summary_sheet)
        if df is None or df.empty:
            return None

        # Find PSP rejection rows (G column = 否 or "No")
        rejections = []
        for idx in range(min(len(df), 30)):
            row = df.iloc[idx]
            g_val = row.iloc[6] if len(row) > 6 else None  # G column - 执行
            h_val = row.iloc[7] if len(row) > 7 else None  # H column - 不执行原因
            c_val = row.iloc[2] if len(row) > 2 else None  # C column - 程序内容

            g_str = str(g_val or "").strip().lower()

            if g_str in ("否", "no", "n") and c_val:
                rejection_text = f"拒绝执行的程序: {c_val[:200]}"
                reason_text = f"拒绝理由: {h_val}" if h_val else ""
                rejections.append((idx, rejection_text, reason_text, str(h_val or "")))

        if not rejections:
            return None

        # Evaluate each rejection with AI
        issues = []
        for idx, prog_text, reason_text, raw_reason in rejections:
            if len(raw_reason.strip()) < 5:
                issues.append((idx, "拒绝理由过于简短或未填写"))
                continue

            prompt = f"""作为审计底稿reviewer，请判断以下PSP拒绝执行理由是否充分：

审计程序: {prog_text[:300]}
拒绝理由: {raw_reason}

判断标准:
- 充分: 理由具体、有业务背景支撑（如"本期无法律费用"需与BKD数据匹配）
- 不充分: 理由笼统、模板化、缺乏具体依据
- 需要进一步验证: 理由看似合理但需结合底稿数据验证

请输出 JSON: {{"has_issue": bool, "reason": "判断理由", "suggestion": "建议"}}
"""

            result = self._ai_judge(prompt)
            if result.get("has_issue"):
                issues.append((idx, f"拒绝理由: {raw_reason[:100]} | AI判断: {result['reason'][:200]}"))

        if issues:
            details = "\n".join([f"  - {i[1]}" for i in issues[:3]])
            return CheckResult(
                rule_id=rule_id,
                rule_name="PSP拒绝执行理由充分性",
                category=CheckCategory.AI_REVIEW,
                severity=Severity.REVIEW,
                location=CellLocation(summary_sheet, issues[0][0] + 1, 7),
                message=f"以下PSP拒绝执行理由需人工复核：\n{details}",
                expected="拒绝执行PSP应有充分、具体的理由",
                actual=f"有{len(issues)}个PSP拒绝理由需关注",
                ai_suggestion="请结合底稿数据和业务背景，评估拒绝理由的充分性",
            )

        return None

    def _check_fluctuation_ai(self, rule_id: str, workbook: AuditWorkbook) -> Optional[CheckResult]:
        """UEXP-AI-001: AI辅助判断异常波动"""
        for sheet_key in ["VC.00 销售费用BKD", "VD.00 管理费用BKD"]:
            sheet_name = None
            for sn in workbook.sheets:
                if sheet_key in sn:
                    sheet_name = sn
                    break
            if not sheet_name:
                continue

            df = workbook.sheets.get(sheet_name)
            if df is None or df.empty:
                continue

            # Find data rows and extract large changes
            large_items = []
            for idx in range(min(len(df), 150)):
                row = df.iloc[idx]
                if len(row) < 12:
                    continue
                d_val = row.iloc[3] if len(row) > 3 else None
                l_val = row.iloc[11] if len(row) > 11 else None
                m_val = row.iloc[12] if len(row) > 12 else None

                if d_val and l_val and m_val:
                    try:
                        la = abs(float(l_val))
                        ma = abs(float(m_val))
                        if la > 500000 and ma > 0.5:  # Significant change: >500K and >50%
                            large_items.append((str(d_val).strip(), float(l_val), float(m_val)))
                    except (ValueError, TypeError):
                        pass

            if large_items:
                items_desc = "\n".join([
                    f"  - {d}: 变动金额{lv:,.0f}, 变动率{mv*100:.1f}%"
                    for d, lv, mv in large_items[:5]
                ])
                prompt = f"""作为审计底稿reviewer，请判断以下费用明细项的波动是否需要进一步调查：

{items_desc}

判断标准：
- 需要调查：变动金额重大且无合理解释、性质异常或与预期不符
- 无需调查：金额虽大但属于正常业务波动（如薪酬根据人数调整）
- 需复核：无法确定

请输出 JSON: {{"has_issue": bool, "reason": "判断理由", "suggestion": "建议"}}
"""
                result = self._ai_judge(prompt)
                if result.get("has_issue"):
                    return CheckResult(
                        rule_id=rule_id,
                        rule_name="BKD异常波动AI辅助判断",
                        category=CheckCategory.AI_REVIEW,
                        severity=Severity.REVIEW,
                        location=CellLocation(sheet_name, 31, 4),
                        message=f"AI识别到{sheet_key}中存在需关注的大幅波动项目：\n{items_desc[:300]}",
                        expected="异常波动应进一步调查并在底稿中记录分析结果",
                        actual=f"AI判断：{result['reason'][:200]}",
                        ai_suggestion=result.get("suggestion", "建议审查这些项目的波动合理性"),
                    )

        return None

    def _check_classification_ai(self, rule_id: str, workbook: AuditWorkbook) -> Optional[CheckResult]:
        """UEXP-AI-002: AI识别费用分类异常"""
        for sheet_key in ["VC.00 销售费用BKD", "VD.00 管理费用BKD"]:
            sheet_name = None
            for sn in workbook.sheets:
                if sheet_key in sn:
                    sheet_name = sn
                    break
            if not sheet_name:
                continue

            df = workbook.sheets.get(sheet_name)
            if df is None or df.empty:
                continue

            # Check for potentially misclassified items
            items = []
            for idx in range(min(len(df), 150)):
                row = df.iloc[idx]
                if len(row) < 4:
                    continue
                d_val = row.iloc[3] if len(row) > 3 else None
                if d_val:
                    name = str(d_val).strip()
                    if any(kw in name for kw in ["包装物", "质保", "佣金", "运输费", "推广费", "修理费"]):
                        items.append(name)

            if items:
                items_desc = "\n".join([f"  - {n}" for n in items[:5]])
                prompt = f"""作为审计底稿reviewer，请判断{sheet_key}中以下费用项目是否存在分类异常风险：

{items_desc}

判断标准：
- 销售费用中的运输费：需判断是否属于合同履约成本（应计入营业成本）
- 销售费用中的包装物：一般应计入主营业务成本
- 销售费用中的质保费用：保证类质保应计入营业成本
- 管理费用中的修理费：需判断受益对象
- 管理费用中的税金：四小税应调整至税金及附加
- 佣金：需判断是否属于合同取得成本

请输出 JSON: {{"has_issue": bool, "reason": "判断理由", "suggestion": "建议"}}
"""
                result = self._ai_judge(prompt)
                if result.get("has_issue"):
                    return CheckResult(
                        rule_id=rule_id,
                        rule_name="费用分类异常AI识别",
                        category=CheckCategory.AI_REVIEW,
                        severity=Severity.REVIEW,
                        location=CellLocation(sheet_name, 31, 4),
                        message=f"AI识别到{sheet_key}中存在需关注分类风险的科目：\n{items_desc[:300]}",
                        expected="费用应按准则要求在正确科目列报",
                        actual=f"AI判断：{result['reason'][:200]}",
                        ai_suggestion=result.get("suggestion", "建议审查这些科目的分类是否正确"),
                    )

        return None

    def _check_legal_fees_ai(self, rule_id: str, workbook: AuditWorkbook) -> Optional[CheckResult]:
        """UEXP-AI-003: AI判断法律费用关注性"""
        sheet_name = None
        for sn in workbook.sheets:
            if "VD.00 管理费用BKD" in sn:
                sheet_name = sn
                break
        if not sheet_name:
            return None

        df = workbook.sheets.get(sheet_name)
        if df is None or df.empty:
            return None

        # Find legal/litigation related items
        legal_items = []
        for idx in range(min(len(df), 150)):
            row = df.iloc[idx]
            if len(row) < 10:
                continue
            d_val = row.iloc[3] if len(row) > 3 else None
            j_val = row.iloc[9] if len(row) > 9 else None
            k_val = row.iloc[10] if len(row) > 10 else None

            if d_val and any(kw in str(d_val) for kw in ["诉讼", "法律", "律师", "法律服务"]):
                try:
                    jn = float(j_val) if j_val else 0
                    kn = float(k_val) if k_val else 0
                except (ValueError, TypeError):
                    jn = 0; kn = 0
                legal_items.append((str(d_val).strip(), jn, kn))

        if not legal_items:
            return None

        # Check if PSP VD.01.03 (复核法律费用) was correctly rejected
        summary_sheet = None
        for sn in workbook.sheets:
            if "汇总" in sn:
                summary_sheet = sn
                break
        psp_skipped = False
        psp_reason = ""
        if summary_sheet:
            sdf = workbook.sheets.get(summary_sheet)
            if sdf is not None:
                for idx in range(min(len(sdf), 30)):
                    row = sdf.iloc[idx]
                    c_val = row.iloc[2] if len(row) > 2 else ""
                    g_val = row.iloc[6] if len(row) > 6 else ""
                    if c_val and "法律费用" in str(c_val):
                        g_str = str(g_val or "").strip().lower()
                        if g_str in ("否", "no", "n"):
                            psp_skipped = True
                            h_val = row.iloc[7] if len(row) > 7 else ""
                            psp_reason = str(h_val or "")

        # If there ARE legal items in BKD but PSP was rejected, flag this
        if psp_skipped and legal_items:
            items_desc = "; ".join([f"{d}(本期{jn:,.0f}, 上期{kn:,.0f})" for d, jn, kn in legal_items])
            return CheckResult(
                rule_id=rule_id,
                rule_name="法律费用关注性判断",
                category=CheckCategory.AI_REVIEW,
                severity=Severity.REVIEW,
                location=CellLocation(sheet_name, 31, 4),
                message=f"管理费用BKD中存在法律/诉讼相关费用（{items_desc}），"
                        f"但汇总页VD.01.03复核法律费用PSP选择不执行。拒绝理由：{psp_reason}",
                expected="若BKD存在法律费用，VD.01.03程序应执行而非拒绝",
                actual="PSP拒绝执行但BKD存在法律费用",
                ai_suggestion="请确认是否存在法律费用相关的审计程序已覆盖，或考虑执行VD.01.03复核法律费用",
            )

        return None
