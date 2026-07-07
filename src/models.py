"""
审计底稿质检Agent - 数据模型定义
定义质检过程中使用的数据结构
"""

from dataclasses import dataclass, field
from typing import List, Optional
from enum import Enum


class Severity(Enum):
    """问题严重程度"""
    HIGH = "高"      # 严重问题，必须修改
    MEDIUM = "中"    # 一般问题，建议修改
    LOW = "低"       # 轻微问题，可选择修改
    REVIEW = "需review"  # 需人工复核（AI判断结果）


class RuleStatus(Enum):
    """规则执行状态"""
    PASSED = "检查通过"          # 规则执行了，没有发现问题
    FAILED = "发现问题"          # 规则执行了，发现了问题
    SKIPPED = "未检查"           # 缺少Sheet/数据，规则未能执行
    AI_DISABLED = "未调用AI"     # AI规则因AI未启用而跳过


class CheckCategory(Enum):
    """检查类别"""
    FORMAT = "格式检查"
    CONSISTENCY = "一致性检查"
    LOGIC = "逻辑检查"
    COMPLETENESS = "完整性检查"
    SCOPE = "底稿范围、程序设计"
    MISSTATEMENT = "错报"
    BASIC_PROC = "基础程序"
    DELIVERY = "交付风险"
    AI_REVIEW = "AI建议人工review"


@dataclass
class CellLocation:
    """单元格位置"""
    sheet: str
    row: int
    col: int

    def to_excel_address(self) -> str:
        """转换为Excel地址格式，如 A1, B2"""
        col_letter = chr(64 + self.col) if self.col <= 26 else chr(64 + (self.col - 1) // 26) + chr(65 + (self.col - 1) % 26)
        return f"{col_letter}{self.row}"

    def __str__(self) -> str:
        return f"{self.sheet}!{self.to_excel_address()}"


@dataclass
class CheckResult:
    """单个检查结果"""
    rule_id: str                    # 规则ID，如 "C001" 或 "AI001"
    rule_name: str                  # 规则名称
    category: CheckCategory         # 检查类别
    severity: Severity              # 严重程度
    location: CellLocation          # 问题所在单元格
    message: str                    # 问题描述
    expected: Optional[str] = None  # 期望值（可选）
    actual: Optional[str] = None    # 实际值（可选）
    ai_suggestion: Optional[str] = None  # AI改进建议（仅AI_REVIEW类别）
    add_comment: bool = True        # 是否在对应单元格添加批注
    show_location: bool = True      # 是否在结果明细中显示具体单元格
    execution_detail: Optional[str] = None  # 执行详情（真实取数、核对过程、判定依据）
    issue_group: Optional[str] = None  # 同一规则内的问题类别；为空时由报告器自动归类
    related_locations: List[CellLocation] = field(default_factory=list)  # 报告合并后包含的全部位置


@dataclass
class RuleExecutionRecord:
    """单条规则的执行记录"""
    rule_id: str                    # 规则ID，如 "AE-001"
    rule_name: str                  # 规则名称
    category: CheckCategory         # 检查类别
    status: RuleStatus              # 执行状态
    skipped_reason: Optional[str] = None  # 未检查的原因
    issues: List['CheckResult'] = field(default_factory=list)  # 该规则发现的问题列表
    execution_detail: Optional[str] = None  # 执行详情（检查通过/未检查时的核对逻辑说明）
    evidence_locations: List[CellLocation] = field(default_factory=list)  # 检查通过时实际核对的位置
    evidence_text: Optional[str] = None  # 无法精确到单元格时的核对范围说明
    evidence_location_count: int = 0  # 实际核对位置数量（可能大于展示的位置数量）


@dataclass
class AuditWorkbook:
    """审计底稿工作簿"""
    file_path: str                  # 文件路径
    sheets: dict = field(default_factory=dict)  # {sheet_name: DataFrame}
    context: dict = field(default_factory=dict)  # 本次质检任务上下文，如TE/SAD/A3基准值

    def get_sheet(self, name: str):
        """获取指定Sheet的数据"""
        return self.sheets.get(name)

    def get_cell_value(self, sheet: str, row: int, col: int):
        """获取指定单元格的值"""
        df = self.sheets.get(sheet)
        if df is not None and row < len(df) and col < len(df.columns):
            return df.iloc[row, col]
        return None


@dataclass
class QualityCheckReport:
    """质检报告"""
    workbook: AuditWorkbook         # 被质检的底稿
    results: List[CheckResult] = field(default_factory=list)  # 检查结果列表

    def add_result(self, result: CheckResult):
        """添加检查结果"""
        self.results.append(result)

    def get_summary(self) -> dict:
        """获取汇总统计"""
        summary = {
            "total": len(self.results),
            "high": sum(1 for r in self.results if r.severity == Severity.HIGH),
            "medium": sum(1 for r in self.results if r.severity == Severity.MEDIUM),
            "low": sum(1 for r in self.results if r.severity == Severity.LOW),
            "review": sum(1 for r in self.results if r.severity == Severity.REVIEW),
        }
        # 按类别统计
        by_category = {}
        for r in self.results:
            cat = r.category.value
            by_category[cat] = by_category.get(cat, 0) + 1
        summary["by_category"] = by_category
        return summary
