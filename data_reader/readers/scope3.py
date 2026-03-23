"""
范围三数据读取模块
=====================

从Excel中提取范围三其他间接排放数据。
"""

import re
from typing import Dict, List, Any

from .base import BaseReader


class Scope3Reader(BaseReader):
    """范围三数据读取器"""

    def extract_all(self) -> Dict[str, Any]:
        """
        提取所有范围三相关数据

        Returns:
            包含范围三数据的字典
        """
        result = {}

        # 提取范围三类别排放数据
        categories_data = self._extract_categories()
        result.update(categories_data)

        # 提取范围三详细数据（从温室气体盘查清册）
        detail_data = self._extract_detail_data()
        result.update(detail_data)

        return result

    def _extract_categories(self) -> Dict[str, Any]:
        """
        提取范围三类别排放数据

        Returns:
            包含范围三各类别排放量的字典
        """
        result = {}
        for i in range(1, 16):
            result[f'scope_3_category_{i}_emissions'] = 0.0

        # 添加范围三类别名称映射（模板需要）
        result['scope_3_category_names'] = {
            1: '外购商品和服务上游排放',
            2: '资本货物上游排放',
            3: '燃料和能源相关活动未包含在范围一和范围二中的上游排放',
            4: '上游运输和配送',
            5: '运营中产生的废弃物',
            6: '员工商务旅行',
            7: '员工通勤',
            8: '上游租赁资产',
            9: '下游运输和配送',
            10: '销售产品的加工',
            11: '销售产品的使用',
            12: '售出产品的加工',
            13: '下游租赁资产',
            14: '特许经营',
            15: '投资'
        }

        # 从"表1温室气体盘查表"中提取范围三各类别数据
        table_sheet = self.find_sheet_by_name('表1', '温室气体盘查表')

        if table_sheet:
            scope3_mapping = {
                '类别1': 'scope_3_category_1_emissions',
                '类别2': 'scope_3_category_2_emissions',
                '类别3': 'scope_3_category_3_emissions',
                '类别4': 'scope_3_category_4_emissions',
                '类别5': 'scope_3_category_5_emissions',
                '类别6': 'scope_3_category_6_emissions',
                '类别7': 'scope_3_category_7_emissions',
                '类别8': 'scope_3_category_8_emissions',
                '类别9': 'scope_3_category_9_emissions',
                '类别10': 'scope_3_category_10_emissions',
                '类别11': 'scope_3_category_11_emissions',
                '类别12': 'scope_3_category_12_emissions',
                '类别13': 'scope_3_category_13_emissions',
                '类别14': 'scope_3_category_14_emissions',
                '类别15': 'scope_3_category_15_emissions',
            }

            for category_key, var_name in scope3_mapping.items():
                for row in table_sheet.iter_rows():
                    a_val = row[0].value if len(row) > 0 else None
                    if a_val and isinstance(a_val, str) and '范围三' in a_val and category_key in a_val:
                        current_row_num = row[0].row
                        if current_row_num + 2 <= table_sheet.max_row:
                            emission_row = table_sheet[current_row_num + 2]
                            b_val = emission_row[1].value if len(emission_row) > 1 else None
                            if b_val and isinstance(b_val, (int, float)):
                                result[var_name] = float(b_val)
                        break

        return result

    def _extract_detail_data(self) -> Dict[str, Any]:
        """
        从温室气体盘查清册中提取范围三详细数据

        Returns:
            包含范围三各类别详细数据的字典
        """
        result = {}
        for i in range(1, 16):
            result[f'scope3_category{i}'] = []

        # 查找表1温室气体盘查表
        table1_sheet = self.find_sheet_by_name('表1', '温室气体盘查表')

        if not table1_sheet:
            print("[范围三详细] 未找到表1温室气体盘查表")
            return result

        print("[范围三详细] 从表1温室气体盘查表提取范围三详细数据")

        # 范围三类别名称映射
        category_names = {
            '类别1': '外购商品和服务上游排放',
            '类别2': '资本货物上游排放',
            '类别3': '燃料和能源相关活动未包含在范围一和范围二中的上游排放',
            '类别4': '上游运输和配送',
            '类别5': '运营中产生的废弃物',
            '类别6': '员工商务旅行',
            '类别7': '员工通勤',
            '类别8': '上游租赁资产',
            '类别9': '下游运输和配送',
            '类别10': '销售产品的加工',
            '类别11': '销售产品的使用',
            '类别12': '售出产品的加工',
            '类别13': '下游租赁资产',
            '类别14': '特许经营',
            '类别15': '投资'
        }

        # 收集详细排放源行（按类别分组）
        category_detail_rows = {}
        for row_idx, row in enumerate(table1_sheet.iter_rows(min_row=5, values_only=True), start=5):
            row_vals = [str(v) if v is not None else '' for v in row[:10]]

            # 检查是否是数据行（第1列是编号）
            if not row_vals[0] or not row_vals[0].strip():
                continue

            # 检查是否是范围三数据（第5列包含"范围三"）
            if len(row_vals) > 4 and '范围三' in row_vals[4]:
                # 提取类别编号
                category_match = re.search(r'类别\s*(\d+)', row_vals[4])
                if category_match:
                    category_num = int(category_match.group(1))
                    if category_num not in category_detail_rows:
                        category_detail_rows[category_num] = []
                    category_detail_rows[category_num].append({
                        'row_idx': row_idx,
                        'data': row
                    })

        # 为每个类别创建数据项
        for category_num in sorted(category_detail_rows.keys()):
            detail_rows = category_detail_rows[category_num]
            category_key = f'类别{category_num}'
            category_var_name = f'scope3_category{category_num}'

            category_items = []
            sub_num = 0
            for row_info in detail_rows:
                row = row_info['data']
                row_vals = [str(v) if v is not None else '' for v in row[:10]]

                # 提取数据
                number = row_vals[0] if len(row_vals) > 0 else ''
                category = row_vals[1] if len(row_vals) > 1 else ''
                emission_source = row_vals[2] if len(row_vals) > 2 else category
                facility = row_vals[3] if len(row_vals) > 3 else ''
                activity_data = self.safe_float(row[5]) if len(row) > 5 else 0
                emission_factor = self.safe_float(row[7]) if len(row) > 7 else 0
                factor_unit = row_vals[8] if len(row_vals) > 8 else ''

                # 计算排放量
                calculated_emission = activity_data * emission_factor

                if calculated_emission > 0.01:
                    sub_num += 1
                    total_formatted = f"{calculated_emission:,.2f}"

                    # 构造排放源名称（包含设施信息）
                    emission_source_name = emission_source
                    if facility and facility != 'None':
                        emission_source_name = f"{emission_source}（{facility}）"

                    category_items.append({
                        'number': f'3.{category_num}.{sub_num}',
                        'emission_source': emission_source_name,
                        'total_green_house_gas_emissions': total_formatted,
                        'CO2_emissions': total_formatted,
                        'CH4_emissions': '0.00',
                        'N2O_emissions': '0.00',
                        'HFCs_emissions': '0.00',
                        'PFCs_emissions': '0.00',
                        'SFs_emissions': '0.00',
                        'NF3_emissions': '0.00'
                    })

            result[category_var_name] = category_items
            print(f"  {category_var_name}: {len(category_items)} 行")

        return result


__all__ = ['Scope3Reader']
