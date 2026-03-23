"""
基础读取器模块
=====================

定义所有专项读取器的基础类。
"""

import openpyxl
from typing import Dict, Any, Optional

from ..protocols import TABLE_PROTOCOLS
from ..fingerprint import TableFingerprint
from ..extractor import ProtocolExtractor


class BaseReader:
    """数据读取器基类"""

    def __init__(self, workbook: openpyxl.Workbook):
        """
        初始化读取器

        Args:
            workbook: openpyxl工作簿对象
        """
        self.workbook = workbook
        self.extractor = ProtocolExtractor()
        self.fingerprint = TableFingerprint()

    def get_protocol_data(self, protocol_name: str) -> list:
        """
        获取特定协议的数据

        Args:
            protocol_name: 协议名称

        Returns:
            提取的数据列表
        """
        # 查找匹配该协议的工作表
        for sheet in self.workbook.worksheets:
            if self.fingerprint.identify(sheet, sheet.title) == protocol_name:
                return self.extractor.extract_from_sheet(sheet, protocol_name)
        return []

    def find_sheet_by_name(self, *name_patterns: str) -> Optional[openpyxl.worksheet.Worksheet]:
        """
        根据名称模式查找工作表

        Args:
            *name_patterns: 名称模式，所有模式都匹配才返回

        Returns:
            匹配的工作表，未找到返回None
        """
        for sheet in self.workbook.worksheets:
            sheet_name = sheet.title
            if all(pattern and pattern in sheet_name for pattern in name_patterns):
                return sheet
        return None

    def safe_float(self, value) -> float:
        """安全地转换为浮点数"""
        try:
            if value is None:
                return 0.0
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def safe_str(self, value) -> str:
        """安全地转换为字符串"""
        if value is None:
            return ''
        return str(value).strip()

    def safe_get_cell(self, row, col_idx):
        """安全获取单元格值"""
        try:
            if col_idx < len(row):
                cell = row[col_idx]
                return cell.value if cell.value is not None else ''
            return ''
        except:
            return ''

    def format_emission(self, val) -> str:
        """格式化排放量数字（保留两位小数）"""
        if val is None:
            return ''
        if val == 0:
            return "0.00"
        try:
            float_value = float(val)
            if float_value == 0:
                return "0.00"
            return f"{float_value:.2f}"
        except (ValueError, TypeError):
            return '0.00'


__all__ = ['BaseReader']
