"""
审计底稿质检Agent - 质检执行引擎
协调和执行所有质检规则（确定性规则 + AI规则）
"""

import traceback
import time as _time
from typing import List, Optional
from .models import (
    AuditWorkbook,
    CellLocation,
    CheckResult,
    RuleExecutionRecord,
    RuleStatus,
    CheckCategory,
    Severity,
)

# 每次模块被重新加载时更新此时间戳（供调试热重载使用）
_MODULE_LOADED_AT = _time.time()
_MODULE_FILE = __file__
from .parser import WorkbookParser
from .rules.format import FormatRules
from .rules.consistency import ConsistencyRules
from .rules.logic import LogicRules
from .rules.completeness import CompletenessRules
from .rules.ai_rules import AIRules
from .ai_prompts import AI_CHECK_TARGETS
from .config import ENABLE_AI_CHECK
from .check_logic import build_execution_detail


class QualityChecker:
    """质检执行引擎"""

    def __init__(self, enable_ai: bool = None):
        self.parser = WorkbookParser()
        self.enable_ai = enable_ai if enable_ai is not None else ENABLE_AI_CHECK

        self.rules = [
            FormatRules(),
            ConsistencyRules(),
            LogicRules(enable_ai=self.enable_ai),
            CompletenessRules(enable_ai=self.enable_ai),
        ]

        # AI规则（纯AI检查点）
        self.ai_rules = None
        if self.enable_ai:
            try:
                self.ai_rules = AIRules()
            except Exception as e:
                print(f"AI规则初始化失败（将跳过AI检查）: {e}")
                self.enable_ai = False

    def _missing_sheet_reason(self, missing: List[str]) -> str:
        """生成缺少Sheet时的未检查原因。"""
        return f"{', '.join(missing)} Sheet不存在，未执行检查"

    def _missing_requirement_reason(self, rule, rule_id: str, missing: List[str]) -> str:
        """生成规则依赖缺失时的未检查原因，允许规则提供业务化说明。"""
        if hasattr(rule, "missing_requirement_reason"):
            return rule.missing_requirement_reason(rule_id, missing)
        return self._missing_sheet_reason(missing)

    def _ai_rule_category(self, rule_id: str) -> CheckCategory:
        category_map = {
            "MT-001": CheckCategory.MISSTATEMENT,
            "UEXP-PSP-002": CheckCategory.SCOPE,
            "UEXP-AI-001": CheckCategory.LOGIC,
            "UEXP-AI-002": CheckCategory.CONSISTENCY,
            "UEXP-AI-003": CheckCategory.SCOPE,
        }
        return category_map.get(rule_id, CheckCategory.AI_REVIEW)

    def _list_ai_rule_defs(self) -> List[dict]:
        """列出纯AI检查规则，即使AI未启用也用于状态报告。"""
        if self.ai_rules:
            return self.ai_rules.list_ai_rule_info()
        return [
            {"rule_id": rule_id, "name": info["name"], "sheets": info["sheets"]}
            for rule_id, info in AI_CHECK_TARGETS.items()
        ]

    def _is_required_sheet_available(self, rule, workbook: AuditWorkbook, rule_id: str, sheet_name: str) -> bool:
        """判断规则依赖的Sheet是否可用，允许个别规则按内容识别非标准命名Sheet。"""
        if hasattr(rule, "is_required_sheet_available"):
            return rule.is_required_sheet_available(workbook, rule_id, sheet_name)
        return self.parser._match_sheet_name(workbook, sheet_name) is not None

    @staticmethod
    def _first_evidence_location(evidence_locations):
        if not evidence_locations:
            return "", ""
        first = evidence_locations[0]
        return first.sheet, first.to_excel_address()

    @staticmethod
    def _should_report_skip_as_review(reason: str) -> bool:
        """规则已执行到业务内容但无法可靠判断时，作为问题项提示人工复核。"""
        reason = str(reason or "")
        if not reason:
            return False
        skip_only_markers = [
            "Sheet不存在",
            "不存在TOD",
            "不适用",
            "缺失",
            "AI未配置",
        ]
        if any(marker in reason for marker in skip_only_markers):
            return False
        review_markers = [
            "已找到",
            "未识别",
            "无法",
            "AI调用失败",
            "未提供",
            "未配置",
            "未记录到实际取数位置",
        ]
        return any(marker in reason for marker in review_markers)

    def _rule_evidence(self, rule, rule_id: str) -> tuple:
        evidence_locations = []
        evidence_text = None
        evidence_location_count = 0
        if hasattr(rule, "get_evidence_locations"):
            evidence_locations = rule.get_evidence_locations(rule_id)
        if hasattr(rule, "get_evidence_text"):
            evidence_text = rule.get_evidence_text(rule_id)
        if hasattr(rule, "get_evidence_location_count"):
            evidence_location_count = rule.get_evidence_location_count(rule_id)
        return evidence_locations, evidence_text, evidence_location_count

    def _fallback_rule_location(self, rule, workbook: AuditWorkbook, rdef: dict) -> CellLocation:
        for sheet_name in rdef.get("required_sheets", []):
            actual = self.parser._match_sheet_name(workbook, sheet_name)
            if actual is not None:
                return CellLocation(actual, 1, 1)
        if workbook.sheets:
            return CellLocation(next(iter(workbook.sheets.keys())), 1, 1)
        return CellLocation("全底稿", 1, 1)

    def _manual_review_issue(
        self,
        rule,
        workbook: AuditWorkbook,
        rdef: dict,
        reason: str,
        evidence_locations: list,
    ) -> CheckResult:
        location = evidence_locations[0] if evidence_locations else self._fallback_rule_location(
            rule, workbook, rdef
        )
        return CheckResult(
            rule_id=rdef["rule_id"],
            rule_name=f"{rdef['rule_name']}需人工复核",
            category=rdef["category"],
            severity=Severity.REVIEW,
            location=location,
            message=(
                f"“{rdef['rule_name']}”已识别到底稿相关内容，但自动规则无法可靠完成判断，"
                "需人工复核。"
            ),
            expected="应能识别完整取数范围并完成自动核对；无法自动判断时需人工复核",
            actual=reason,
            ai_suggestion="请审计师或reviewer查看对应底稿区域，确认该检查点是否存在遗漏或异常。",
        )

    def check(self, file_path: str, context: Optional[dict] = None) -> tuple:
        # 解析底稿
        workbook = self.parser.parse(file_path)
        workbook.context = context or {}

        # 验证结构
        missing_sheets = self.parser.validate_structure(workbook)
        if missing_sheets:
            print(f"警告：底稿缺少以下标准Sheet: {missing_sheets}")

        # 执行确定性规则检查
        all_results: List[CheckResult] = []
        execution_records: List[RuleExecutionRecord] = []

        for rule in self.rules:
            rule_defs = rule.list_rule_ids()
            skipped_rule_ids = set()

            # 记录未检查的规则（缺Sheet→跳过，与check()是否成功无关）
            for rdef in rule_defs:
                required = rdef.get("required_sheets", [])
                missing = [
                    s for s in required
                    if not self._is_required_sheet_available(rule, workbook, rdef["rule_id"], s)
                ]
                if missing:
                    skipped_rule_ids.add(rdef["rule_id"])
                    missing_reason = self._missing_requirement_reason(
                        rule, rdef["rule_id"], missing
                    )
                    if self._should_report_skip_as_review(missing_reason):
                        review_issue = self._manual_review_issue(
                            rule,
                            workbook,
                            rdef,
                            missing_reason,
                            [],
                        )
                        review_issue.execution_detail = build_execution_detail(
                            rdef["rule_id"],
                            message=review_issue.message,
                            expected=review_issue.expected or "",
                            actual=review_issue.actual or "",
                            sheet=review_issue.location.sheet,
                            cell=review_issue.location.to_excel_address(),
                            context=getattr(workbook, 'context', {}),
                            status="FAILED",
                        )
                        all_results.append(review_issue)
                        execution_records.append(RuleExecutionRecord(
                            rule_id=rdef["rule_id"],
                            rule_name=rdef["rule_name"],
                            category=rdef["category"],
                            status=RuleStatus.FAILED,
                            skipped_reason=missing_reason,
                            issues=[review_issue],
                        ))
                        continue
                    execution_records.append(RuleExecutionRecord(
                        rule_id=rdef["rule_id"],
                        rule_name=rdef["rule_name"],
                        category=rdef["category"],
                        status=RuleStatus.SKIPPED,
                        skipped_reason=missing_reason
                    ))

            # 执行检查（只对check()调用做异常保护，记录生成逻辑始终运行）
            check_error = None
            try:
                results = rule.check(workbook)
                results = [r for r in results if r.rule_id not in skipped_rule_ids]
                # 为每个问题项构建 execution_detail
                ctx = getattr(workbook, 'context', {})
                for r in results:
                    if r.execution_detail is None:
                        r.execution_detail = build_execution_detail(
                            r.rule_id,
                            message=r.message,
                            expected=r.expected or "",
                            actual=r.actual or "",
                            sheet=r.location.sheet if r.show_location else "",
                            cell=r.location.to_excel_address() if r.show_location else "",
                            context=ctx,
                            status="FAILED"
                        )
                all_results.extend(results)
            except Exception as e:
                print(f"规则模块 {rule.rule_name} 执行出错: {e}")
                traceback.print_exc()
                results = []
                check_error = str(e)

            # 按 rule_id 分组结果，生成执行记录
            results_by_rule = {}
            for r in results:
                results_by_rule.setdefault(r.rule_id, []).append(r)
            runtime_errors = getattr(rule, "runtime_errors", {})
            runtime_skips = getattr(rule, "runtime_skips", {})

            for rdef in rule_defs:
                if rdef["rule_id"] in skipped_rule_ids:
                    continue
                issues = results_by_rule.get(rdef["rule_id"], [])
                # 模块级异常：所有未跳过的规则标记为执行失败
                if check_error is not None:
                    execution_records.append(RuleExecutionRecord(
                        rule_id=rdef["rule_id"],
                        rule_name=rdef["rule_name"],
                        category=rdef["category"],
                        status=RuleStatus.FAILED,
                        skipped_reason=f"规则模块执行异常，请联系管理员排查: {check_error}",
                        issues=issues
                    ))
                    continue
                if rdef["rule_id"] in runtime_errors:
                    execution_records.append(RuleExecutionRecord(
                        rule_id=rdef["rule_id"],
                        rule_name=rdef["rule_name"],
                        category=rdef["category"],
                        status=RuleStatus.FAILED,
                        skipped_reason=(
                            "子规则执行异常，请联系管理员排查: "
                            f"{runtime_errors[rdef['rule_id']]}"
                        ),
                        issues=issues
                    ))
                    continue
                if rdef["rule_id"] in runtime_skips and not issues:
                    runtime_skip_reason = runtime_skips[rdef["rule_id"]]
                    evidence_locations, evidence_text, evidence_location_count = self._rule_evidence(
                        rule, rdef["rule_id"]
                    )
                    if self._should_report_skip_as_review(runtime_skip_reason):
                        review_issue = self._manual_review_issue(
                            rule,
                            workbook,
                            rdef,
                            runtime_skip_reason,
                            evidence_locations,
                        )
                        review_issue.execution_detail = build_execution_detail(
                            rdef["rule_id"],
                            message=review_issue.message,
                            expected=review_issue.expected or "",
                            actual=review_issue.actual or "",
                            sheet=review_issue.location.sheet,
                            cell=review_issue.location.to_excel_address(),
                            context=getattr(workbook, 'context', {}),
                            status="FAILED",
                        )
                        all_results.append(review_issue)
                        execution_records.append(RuleExecutionRecord(
                            rule_id=rdef["rule_id"],
                            rule_name=rdef["rule_name"],
                            category=rdef["category"],
                            status=RuleStatus.FAILED,
                            skipped_reason=runtime_skip_reason,
                            issues=[review_issue],
                            evidence_locations=evidence_locations,
                            evidence_text=evidence_text,
                            evidence_location_count=evidence_location_count,
                        ))
                        continue
                    execution_records.append(RuleExecutionRecord(
                        rule_id=rdef["rule_id"],
                        rule_name=rdef["rule_name"],
                        category=rdef["category"],
                        status=RuleStatus.SKIPPED,
                        skipped_reason=runtime_skip_reason,
                        issues=[],
                        evidence_locations=evidence_locations,
                        evidence_text=evidence_text,
                        evidence_location_count=evidence_location_count,
                    ))
                    continue
                runtime_skip_reason = runtime_skips.get(rdef["rule_id"])
                # 检查是否为需要AI的规则但AI未配置
                if hasattr(rule, "needs_ai_for_workbook"):
                    needs_ai = rule.needs_ai_for_workbook(rdef["rule_id"], workbook)
                else:
                    needs_ai = hasattr(rule, 'needs_ai') and rule.needs_ai(rdef["rule_id"])
                ai_disabled = needs_ai and getattr(rule, "ai_client", None) is None
                if ai_disabled and issues:
                    status = RuleStatus.FAILED
                    skipped_reason = "AI未配置，已执行Python部分，AI部分未执行"
                elif ai_disabled:
                    status = RuleStatus.AI_DISABLED
                    skipped_reason = "AI未配置，无法执行AI判断"
                else:
                    status = RuleStatus.FAILED if issues else RuleStatus.PASSED
                    skipped_reason = runtime_skip_reason if issues and runtime_skip_reason else None
                # 为检查通过项构建 execution_detail
                passed_detail = None
                evidence_locations = []
                evidence_text = None
                evidence_location_count = 0
                if status == RuleStatus.PASSED:
                    try:
                        evidence_locations, evidence_text, evidence_location_count = self._rule_evidence(
                            rule, rdef["rule_id"]
                        )
                        if not evidence_locations and not evidence_text:
                            status = RuleStatus.SKIPPED
                            skipped_reason = "未记录到实际取数位置，未确认规则已完成有效核对"
                        else:
                            evidence_sheet, evidence_cell = self._first_evidence_location(
                                evidence_locations
                            )
                            passed_detail = build_execution_detail(
                                rdef["rule_id"],
                                context=getattr(workbook, 'context', {}),
                                status="PASSED",
                                sheet=evidence_sheet,
                                cell=evidence_cell,
                            )
                    except Exception:
                        status = RuleStatus.SKIPPED
                        skipped_reason = "生成通过项取数证据失败，未确认规则已完成有效核对"
                execution_records.append(RuleExecutionRecord(
                    rule_id=rdef["rule_id"],
                    rule_name=rdef["rule_name"],
                    category=rdef["category"],
                    status=status,
                    skipped_reason=skipped_reason,
                    issues=issues,
                    execution_detail=passed_detail,
                    evidence_locations=evidence_locations,
                    evidence_text=evidence_text,
                    evidence_location_count=evidence_location_count,
                ))

        # 执行AI规则检查
        ai_rule_defs = self._list_ai_rule_defs()
        skipped_ai_ids = set()

        # 缺少依赖Sheet优先归类为"未检查"，不归类为"通过/问题/未调用AI"
        for adef in ai_rule_defs:
            missing = [
                s for s in adef["sheets"]
                if self.parser._match_sheet_name(workbook, s) is None
            ]
            if missing:
                skipped_ai_ids.add(adef["rule_id"])
                execution_records.append(RuleExecutionRecord(
                    rule_id=adef["rule_id"],
                    rule_name=adef["name"],
                    category=self._ai_rule_category(adef["rule_id"]),
                    status=RuleStatus.SKIPPED,
                    skipped_reason=self._missing_sheet_reason(missing)
                ))

        ai_enabled = self.enable_ai and self.ai_rules and self.ai_rules.is_enabled()

        if not ai_enabled:
            # AI未启用；已因缺Sheet跳过的AI规则不再重复标记
            for adef in ai_rule_defs:
                if adef["rule_id"] in skipped_ai_ids:
                    continue
                execution_records.append(RuleExecutionRecord(
                    rule_id=adef["rule_id"],
                    rule_name=adef["name"],
                    category=self._ai_rule_category(adef["rule_id"]),
                    status=RuleStatus.AI_DISABLED,
                    skipped_reason="AI功能未启用或API密钥未配置"
                ))
        else:
            try:
                ai_results = self.ai_rules.check(workbook)
                ai_results = [r for r in ai_results if r.rule_id not in skipped_ai_ids]
                # 为AI问题项构建 execution_detail
                ctx = getattr(workbook, 'context', {})
                for r in ai_results:
                    if r.execution_detail is None:
                        r.execution_detail = build_execution_detail(
                            r.rule_id,
                            message=r.message,
                            expected=r.expected or "",
                            actual=r.actual or "",
                            sheet=r.location.sheet if r.show_location else "",
                            cell=r.location.to_excel_address() if r.show_location else "",
                            context=ctx,
                            status="FAILED"
                        )
                all_results.extend(ai_results)
                print(f"AI检查完成，发现 {len(ai_results)} 个需review的问题")

                results_by_rule = {}
                for r in ai_results:
                    results_by_rule.setdefault(r.rule_id, []).append(r)
                runtime_skips = getattr(self.ai_rules, "runtime_skips", {})

                for adef in ai_rule_defs:
                    if adef["rule_id"] in skipped_ai_ids:
                        continue
                    issues = results_by_rule.get(adef["rule_id"], [])
                    runtime_skip_reason = runtime_skips.get(adef["rule_id"])
                    if runtime_skip_reason and not issues:
                        execution_records.append(RuleExecutionRecord(
                            rule_id=adef["rule_id"],
                            rule_name=adef["name"],
                            category=self._ai_rule_category(adef["rule_id"]),
                            status=RuleStatus.SKIPPED,
                            skipped_reason=runtime_skip_reason,
                            issues=[],
                        ))
                        continue
                    ai_passed_detail = None
                    evidence_locations = []
                    evidence_text = None
                    evidence_location_count = 0
                    if not issues:
                        try:
                            if hasattr(self.ai_rules, "get_evidence_locations"):
                                evidence_locations = self.ai_rules.get_evidence_locations(adef["rule_id"])
                            if hasattr(self.ai_rules, "get_evidence_text"):
                                evidence_text = self.ai_rules.get_evidence_text(adef["rule_id"])
                            if hasattr(self.ai_rules, "get_evidence_location_count"):
                                evidence_location_count = self.ai_rules.get_evidence_location_count(adef["rule_id"])
                            if not evidence_locations and not evidence_text:
                                runtime_skip_reason = "未记录到实际取数位置，未确认AI规则已完成有效核对"
                            else:
                                evidence_sheet, evidence_cell = self._first_evidence_location(
                                    evidence_locations
                                )
                                ai_passed_detail = build_execution_detail(
                                    adef["rule_id"],
                                    context=ctx,
                                    status="PASSED",
                                    sheet=evidence_sheet,
                                    cell=evidence_cell,
                                )
                        except Exception:
                            runtime_skip_reason = "生成AI通过项取数证据失败，未确认规则已完成有效核对"
                    execution_records.append(RuleExecutionRecord(
                        rule_id=adef["rule_id"],
                        rule_name=adef["name"],
                        category=self._ai_rule_category(adef["rule_id"]),
                        status=RuleStatus.FAILED if issues else (
                            RuleStatus.SKIPPED if runtime_skip_reason else RuleStatus.PASSED
                        ),
                        skipped_reason=runtime_skip_reason if runtime_skip_reason else None,
                        issues=issues,
                        execution_detail=ai_passed_detail,
                        evidence_locations=evidence_locations,
                        evidence_text=evidence_text,
                        evidence_location_count=evidence_location_count,
                    ))
            except Exception as e:
                print(f"AI规则执行出错（已跳过）: {e}")

        return workbook, all_results, execution_records
