"""
工具函数模块
=====================

通用工具函数，可在整个包中使用。
"""

import re
from datetime import datetime
from typing import Any, Dict


def excel_date_to_string(date_value):
    """将Excel日期序列号转换为 'YYYY年MM月DD日' 格式"""
    if date_value is None:
        return None
    if isinstance(date_value, str):
        return date_value
    try:
        # Excel日期基准是1899-12-30
        delta = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_value) - 2)
        return delta.strftime('%Y年%m月%d日')
    except (ValueError, TypeError):
        return str(date_value)


def safe_str(value) -> str:
    """安全地转换为字符串"""
    if value is None:
        return ''
    return str(value).strip()


def safe_float(value) -> float:
    """安全地转换为浮点数"""
    try:
        if value is None:
            return 0.0
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def safe_get_cell(row, col_idx):
    """安全获取单元格值"""
    try:
        if col_idx < len(row):
            cell = row[col_idx]
            return cell.value if cell.value is not None else ''
        return ''
    except:
        return ''


def clean_multiline_text(value: str) -> str:
    """清理多行文本，去除多余空白"""
    if value is None:
        return ''
    if isinstance(value, str):
        value = re.sub(r'[\n\r]+', ' ', value)
        value = re.sub(r'\s+', ' ', value).strip()
    return value


__all__ = [
    'excel_date_to_string',
    'safe_str',
    'safe_float',
    'safe_get_cell',
    'clean_multiline_text',
]
