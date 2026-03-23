"""
表格指纹识别模块
=====================

负责识别Excel工作表的表格类型。
"""

import openpyxl
from typing import Optional, Dict

from .protocols import TABLE_PROTOCOLS, _PROTOCOL_ORDER
from .config import TableProtocol


class TableFingerprint:
    """表格指纹识别器"""

    def __init__(self, protocols: Dict[str, TableProtocol] = None):
        self.protocols = protocols or TABLE_PROTOCOLS

    def identify(self, sheet: openpyxl.worksheet.Worksheet,
                 sheet_name: str = None,
                 check_rows: int = 20) -> Optional[str]:
        """
        识别工作表的表格类型

        按照协议优先级顺序进行匹配，找到第一个匹配的协议就返回

        Args:
            sheet: openpyxl工作表对象
            sheet_name: 工作表名称
            check_rows: 检查前N行

        Returns:
            匹配的协议名称，未匹配返回None
        """
        # 收集表格的唯一字符串值
        unique_strings = self._extract_unique_strings(sheet, check_rows)

        actual_sheet_name = sheet_name or sheet.title

        # 先尝试精确的工作表名称匹配（优先级最高）
        for protocol_name in _PROTOCOL_ORDER:
            if protocol_name not in self.protocols:
                continue

            protocol = self.protocols[protocol_name]

            # 检查工作表名称模式
            if protocol.sheet_name_patterns:
                # 检查所有模式是否都匹配
                all_patterns_match = all(pattern and pattern in actual_sheet_name for pattern in protocol.sheet_name_patterns)
                if all_patterns_match:
                    print(f"[表格识别] '{actual_sheet_name}' 通过名称匹配到 {protocol.name} (patterns: {protocol.sheet_name_patterns})")
                    return protocol_name

        # 如果没有名称匹配，则按关键词匹配
        for protocol_name in _PROTOCOL_ORDER:
            if protocol_name not in self.protocols:
                continue

            protocol = self.protocols[protocol_name]

            # 跳过有名称模式但未匹配的协议（因为它们在前面已经检查过了）
            if protocol.sheet_name_patterns:
                continue

            # 检查必需关键词是否全部匹配
            if not protocol.required_keywords.issubset(unique_strings):
                continue

            # 计算可选关键词匹配度
            optional_matched = len(protocol.optional_keywords & unique_strings)
            optional_total = len(protocol.optional_keywords) if protocol.optional_keywords else 1
            match_score = optional_matched / optional_total if optional_total > 0 else 1.0

            # 检查是否达到最小匹配度
            if match_score >= protocol.min_match_ratio:
                print(f"[表格识别] {sheet_name or sheet.title} 匹配到 {protocol.name} "
                      f"(必需关键词: {len(protocol.required_keywords)}/{len(protocol.required_keywords)}, "
                      f"可选关键词: {optional_matched}/{optional_total})")
                return protocol_name

        if sheet_name:
            print(f"[表格识别] {sheet_name} 未匹配到已知协议类型")
        return None

    def _extract_unique_strings(self, sheet: openpyxl.worksheet.Worksheet,
                               check_rows: int) -> set:
        """提取表格中的唯一字符串"""
        unique_strings = set()
        for row_idx in range(1, min(check_rows + 1, sheet.max_row + 1)):
            for cell in sheet[row_idx]:
                if cell.value and isinstance(cell.value, str):
                    cleaned = str(cell.value).strip()
                    if cleaned:
                        unique_strings.add(cleaned)
        return unique_strings

    def _calculate_match_score(self, unique_strings: set,
                              protocol: TableProtocol) -> float:
        """计算协议匹配度"""
        # 检查必需关键词
        if not protocol.required_keywords.issubset(unique_strings):
            return 0.0

        # 计算可选关键词匹配度
        optional_matched = len(protocol.optional_keywords & unique_strings)
        optional_total = len(protocol.optional_keywords) if protocol.optional_keywords else 1

        match_ratio = optional_matched / optional_total if optional_total > 0 else 1.0
        return match_ratio


__all__ = ['TableFingerprint']
