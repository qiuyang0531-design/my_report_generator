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

    def is_error_value(self, value) -> bool:
        """检查是否为 Excel 错误值"""
        if value is None:
            return False
        v_str = str(value).strip().upper()
        return v_str in ['#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!']

    def natural_sort_key(self, s: str):
        """用于编号（如 1.1, 1.10, 3.2.3）的自然排序键"""
        import re
        if not s:
            return []
        # 将字符串拆分为数字和非数字部分，并将数字部分转换为整数
        return [int(text) if text.isdigit() else text.lower()
                for text in re.split('([0-9]+)', str(s))]

    def safe_float(self, value) -> float:
        """安全地转换为浮点数，处理 Excel 错误值"""
        try:
            if value is None:
                return 0.0
            
            # 处理常见的 Excel 错误字符串
            if isinstance(value, str):
                v_str = value.strip().upper()
                if v_str in ['#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!']:
                    return 0.0
                # 去除逗号
                v_str = v_str.replace(',', '')
                if not v_str:
                    return 0.0
                return float(v_str)
                
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def safe_str(self, value) -> str:
        """安全地转换为字符串，处理 Excel 错误值"""
        if value is None:
            return ''
        
        v_str = str(value).strip()
        # 处理常见的 Excel 错误字符串
        if v_str.upper() in ['#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!']:
            return ''
            
        return v_str

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
