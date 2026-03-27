"""
基本信息读取模块
=====================

从Excel中提取基本信息数据。
"""

import re
from typing import Dict, Any

from .base import BaseReader
from ..utils import excel_date_to_string, clean_multiline_text


class BasicInfoReader(BaseReader):
    """基本信息读取器"""

    def extract(self) -> Dict[str, Any]:
        """
        提取基本信息

        Returns:
            包含基本信息的字典
        """
        result = {
            'company_name': None,
            'company_profile': None,
            'legal_person': None,
            'registered_address': None,
            'date_of_establishment': None,
            'registered_capital': None,
            'Unified_Social_Credit_Identifier': None,
            'deadline': None,
            'evaluation_level': None,
            'evaluation_score': None,
            'scope_of_business': None,
            'rule_file': None,
            'GWP_Value_Reference_Document': None,
            'document_number': None,
            'posted_time': None,
            'production_address': None,
            'reporting_period': None,
            'report_year': '2024',
            'next_year': '2025',  # 明年
            'producer': None,
            'auditor': None,
            'approver': None,
        }

        # 查找基本信息表
        basic_info_sheet = self.find_sheet_by_name('基本信息')

        if basic_info_sheet:
            print(f"[基本信息] 找到基本信息表: {basic_info_sheet.title}")
            # 读取基本信息：第2列是属性代码(key)，第3列是值(value)
            for row in basic_info_sheet.iter_rows(min_row=2, max_row=50, values_only=True):
                if len(row) >= 3 and row[1] and row[2]:
                    key = str(row[1]).strip()  # 第2列：属性代码
                    value = row[2]  # 第3列：值

                    # 映射到标准字段
                    if 'company_name' in key or '组织名称' in key:
                        result['company_name'] = str(value).strip() if value else None
                    elif key == 'company_profile':
                        result['company_profile'] = clean_multiline_text(value)
                    elif key == 'scope_of_business':
                        result['scope_of_business'] = clean_multiline_text(value)
                    elif key == 'legal_person':
                        result['legal_person'] = str(value).strip() if value else None
                    elif key == 'registered_address':
                        result['registered_address'] = str(value).strip() if value else None
                    elif key == 'production_address':
                        result['production_address'] = str(value).strip() if value else None
                    elif key == 'date_of_establishment':
                        result['date_of_establishment'] = excel_date_to_string(value)
                    elif key == 'posted_time':
                        result['posted_time'] = excel_date_to_string(value)
                    elif key == 'deadline':
                        result['deadline'] = excel_date_to_string(value)
                    elif key == 'registered_capital':
                        result['registered_capital'] = str(value).strip() if value else None
                    elif key == 'Unified_Social_Credit_Identifier':
                        result['Unified_Social_Credit_Identifier'] = str(value).strip() if value else None
                    elif key == 'reporting_period':
                        result['reporting_period'] = str(value).strip() if value else None
                        # 从周期中提取年份
                        year_match = re.search(r'(\d{4})', str(value))
                        result['report_year'] = year_match.group(1) if year_match else '2024'
                        # 计算明年
                        try:
                            result['next_year'] = str(int(result['report_year']) + 1)
                        except (ValueError, TypeError):
                            result['next_year'] = '2025'
                    elif key == 'document_number':
                        result['document_number'] = str(value).strip() if value else None
                    elif key in result:
                        result[key] = value

            print(f"[基本信息] 公司名称: {result.get('company_name')}")
        else:
            print("[基本信息] 未找到基本信息表，尝试从温室气体盘查清册提取...")
            # 尝试从温室气体盘查清册提取
            inventory_sheet = self.find_sheet_by_name('盘查清册', '清册')
            if inventory_sheet:
                for row in inventory_sheet.iter_rows(max_row=20, values_only=True):
                    if len(row) >= 3:
                        if row[1] == '组织名称：' and row[2]:
                            result['company_name'] = row[2]
                        elif row[1] == '组织地址：' and row[2]:
                            result['registered_address'] = row[2]
                            result['production_address'] = row[2]
                        elif row[1] == '盘查覆盖周期:' and row[2]:
                            result['reporting_period'] = row[2]
                            year_match = re.search(r'(\d{4})', str(row[2]))
                            result['report_year'] = year_match.group(1) if year_match else '2024'
                            # 计算明年
                            try:
                                result['next_year'] = str(int(result['report_year']) + 1)
                            except (ValueError, TypeError):
                                result['next_year'] = '2025'

        # 将None值转换为空字符串（避免Jinja2渲染为"None"）
        for key in result:
            if result[key] is None:
                result[key] = ''

        return result


__all__ = ['BasicInfoReader']
