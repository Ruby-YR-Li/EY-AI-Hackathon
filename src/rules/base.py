"""
审计底稿质检Agent - 规则基类
定义所有质检规则的基础接口
"""

from abc import ABC, abstractmethod
from contextlib import contextmanager
from typing import List

from ..models import CellLocation


class BaseRule(ABC):
    """规则基类"""

    MAX_EVIDENCE_LOCATIONS_PER_RULE = 50

    def __init__(self, rule_id: str, rule_name: str):
        self.rule_id = rule_id
        self.rule_name = rule_name
        self.runtime_errors = {}
        self.runtime_skips = {}
        self.evidence_locations = {}
        self.evidence_location_counts = {}
        self._evidence_location_keys = {}
        self.evidence_texts = {}
        self._evidence_overflow_rules = set()
        self._current_evidence_rule_id = None

    def reset_runtime_state(self):
        """清空本次工作簿执行过程中记录的子规则异常和动态跳过原因。"""
        self.runtime_errors = {}
        self.runtime_skips = {}
        self.evidence_locations = {}
        self.evidence_location_counts = {}
        self._evidence_location_keys = {}
        self.evidence_texts = {}
        self._evidence_overflow_rules = set()
        self._current_evidence_rule_id = None

    def mark_rule_error(self, rule_id: str, reason: str):
        """记录子规则执行异常，避免执行器把空结果误判为检查通过。"""
        self.runtime_errors[rule_id] = reason

    def mark_rule_skipped(self, rule_id: str, reason: str):
        """记录因数据无法识别等原因未能完成的检查。"""
        self.runtime_skips[rule_id] = reason

    def attach_parser(self, parser):
        """让WorkbookParser把实际取数位置回写到当前规则。"""
        parser.evidence_recorder = self._record_parser_evidence
        return parser

    @contextmanager
    def evidence_context(self, rule_id: str):
        """标记当前正在执行的子规则，用于自动记录取数位置。"""
        previous = self._current_evidence_rule_id
        self._current_evidence_rule_id = rule_id
        try:
            yield
        finally:
            self._current_evidence_rule_id = previous

    def record_evidence_location(self, rule_id: str, location: CellLocation):
        """记录某条规则实际核对过的单元格。"""
        if not location:
            return
        locations = self.evidence_locations.setdefault(rule_id, [])
        keys = self._evidence_location_keys.setdefault(rule_id, set())
        key = (location.sheet, location.row, location.col)
        if key not in keys:
            keys.add(key)
            self.evidence_location_counts[rule_id] = (
                self.evidence_location_counts.get(rule_id, 0) + 1
            )
            if len(locations) >= self.MAX_EVIDENCE_LOCATIONS_PER_RULE:
                self._evidence_overflow_rules.add(rule_id)
                return
            locations.append(location)

    def record_evidence_text(self, rule_id: str, text: str):
        """记录无法精确到单元格时的核对范围/外部输入。"""
        text = str(text or "").strip()
        if not text:
            return
        texts = self.evidence_texts.setdefault(rule_id, [])
        if text not in texts:
            texts.append(text)

    def get_evidence_locations(self, rule_id: str) -> List[CellLocation]:
        return list(self.evidence_locations.get(rule_id, []))

    def get_evidence_location_count(self, rule_id: str) -> int:
        return self.evidence_location_counts.get(
            rule_id,
            len(self.evidence_locations.get(rule_id, [])),
        )

    def get_evidence_text(self, rule_id: str) -> str:
        return "；".join(self.evidence_texts.get(rule_id, []))

    def _record_parser_evidence(self, sheet: str, row: int = None, col: int = None, text: str = ""):
        rule_id = self._current_evidence_rule_id
        if not rule_id:
            return
        if sheet and row is not None and col is not None:
            self.record_evidence_location(rule_id, CellLocation(sheet, row + 1, col + 1))
        elif text:
            self.record_evidence_text(rule_id, text)
        elif sheet:
            self.record_evidence_text(rule_id, sheet)

    @abstractmethod
    def check(self, workbook) -> List:
        """执行检查，返回检查结果列表"""
        pass

    def list_rule_ids(self) -> List[dict]:
        """
        列出该规则类包含的所有子规则ID和名称。
        返回: [{"rule_id": "AE-001", "rule_name": "...", "category": ..., "required_sheets": [...]}, ...]
        子类应覆盖此方法。
        """
        return []
