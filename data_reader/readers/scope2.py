"""
范围二数据读取模块
=====================

从Excel中提取范围二外购能源间接排放数据。
"""

import openpyxl
from typing import Dict, List, Any

from .base import BaseReader


class Scope2Reader(BaseReader):
    """范围二数据读取器"""

    def extract_all(self) -> Dict[str, Any]:
        """
        提取所有范围二相关数据

        Returns:
            包含范围二数据的字典
        """
        result = {
            'scope_1_emissions': None,
            'scope_2_location_based_emissions': None,
            'scope_2_market_based_emissions': None,
            'scope_3_emissions': None,
            'scope_2_location': None,
            'scope_2_market': None,
            'total_emission_location': None,
            'total_emission_market': None,
        }

        # 从表1温室气体盘查表提取范围一二三排放汇总数据
        summary_data = self._extract_summary_from_table1()
        result.update(summary_data)

        # 从温室气体盘查表提取范围二输入能源的间接排放清册数据
        items_data = self._extract_scope2_items()
        result.update(items_data)

        return result

    def _extract_summary_from_table1(self) -> Dict[str, Any]:
        """
        从表1温室气体盘查表提取范围一二三排放汇总数据

        Returns:
            包含范围一二三排放汇总的字典
        """
        result = {
            'scope_1_emissions': None,
            'scope_2_location_based_emissions': None,
            'scope_2_market_based_emissions': None,
            'scope_3_emissions': None,
            'scope_2_location': None,
            'scope_2_market': None,
            'total_emission_location': None,
            'total_emission_market': None,
        }

        # 查找表1温室气体盘查表
        table_sheet = self.find_sheet_by_name('表1', '温室气体盘查表')

        if table_sheet:
            # 动态查找总排放量汇总行
            for row in table_sheet.iter_rows(values_only=True):
                a_val = row[0] if len(row) > 0 else None
                b_val = row[1] if len(row) > 1 else None
                c_val = row[2] if len(row) > 2 else None
                d_val = row[3] if len(row) > 3 else None
                e_val = row[4] if len(row) > 4 else None

                if a_val and isinstance(a_val, str) and '排放量' in a_val:
                    if isinstance(b_val, (int, float)) and isinstance(c_val, (int, float)) and isinstance(d_val, (int, float)):
                        result['scope_1_emissions'] = float(b_val)
                        result['scope_2_location_based_emissions'] = float(c_val)
                        result['scope_3_emissions'] = float(d_val)
                        result['scope_2_location'] = float(c_val)
                        if isinstance(e_val, (int, float)):
                            result['total_emission_location'] = float(e_val)
                    break

            # 动态查找范围二基于市场的排放量
            for row in table_sheet.iter_rows():
                e_val = row[4].value if len(row) > 4 else None
                c_val = row[2].value if len(row) > 2 else None
                if e_val and isinstance(e_val, str) and '基于市场' in e_val:
                    if c_val and isinstance(c_val, (int, float)):
                        result['scope_2_market_based_emissions'] = float(c_val)
                        result['scope_2_market'] = float(c_val)
                        print(f"[范围二] 找到基于市场排放量: {float(c_val)}")
                    break

            # 计算总排放量（基于市场）
            result['total_emission_market'] = (
                result.get('scope_1_emissions', 0) +
                result.get('scope_2_market_based_emissions', 0) +
                result.get('scope_3_emissions', 0)
            )

        return result

    def _extract_scope2_items(self) -> Dict[str, Any]:
        """
        从温室气体盘查表中提取范围二输入能源的间接排放清册数据

        Returns:
            包含范围二排放明细的字典
        """
        result = {'scope2_items': []}

        # 查找温室气体盘查表
        pandata_sheet = self.find_sheet_by_name('盘查表')

        if pandata_sheet:
            print(f"[范围二] 找到温室气体盘查表: {pandata_sheet.title}")
            scope2_items = []

            # 查找包含"汇总"和"外购电力"的行
            location_total = None
            market_total = None

            # 尝试匹配包含"汇总"或"Total"的行
            total_keywords = ['汇总', '总计', 'Total', 'TOTAL']

            for row_idx, row in enumerate(pandata_sheet.iter_rows(min_row=1, values_only=True), start=1):
                first_col = str(row[0]).strip() if row[0] else ''
                second_col = str(row[1]).strip() if len(row) > 1 else ''
                fourth_col = str(row[4]).strip() if len(row) > 4 else ''

                # 检查是否是汇总行
                is_total_row = any(keyword in first_col or first_col == keyword for keyword in total_keywords)
                has_electricity = '外购' in second_col or '电力' in second_col

                if is_total_row and has_electricity:
                    # 检查第四列来判断是基于位置还是基于市场
                    is_location = '位置' in fourth_col or 'Location' in fourth_col or 'location' in fourth_col
                    is_market = '市场' in fourth_col or 'Market' in fourth_col or 'market' in fourth_col

                    if is_location:
                        location_total = row
                        print(f"  找到外购电力（基于位置）汇总: Row {row_idx}")
                    elif is_market:
                        market_total = row
                        print(f"  找到外购电力（基于市场）汇总: Row {row_idx}")

            # 构建两行数据
            if location_total:
                total = self.safe_float(location_total[2]) if len(location_total) > 2 else 0
                co2 = self.safe_float(location_total[3]) if len(location_total) > 3 else total

                total_formatted = f"{total:,.2f}" if total > 0 else "0.00"
                co2_formatted = f"{co2:,.2f}" if co2 > 0 else "0.00"

                scope2_items.append({
                    'number': '2.1',
                    'emission_source': '外购电力（基于位置）',
                    'total_green_house_gas_emissions': total_formatted,
                    'CO2_emissions': co2_formatted,
                    'CH4_emissions': '0.00',
                    'N2O_emissions': '0.00',
                    'HFCs_emissions': '0.00',
                    'PFCs_emissions': '0.00',
                    'SFs_emissions': '0.00',
                    'NF3_emissions': '0.00'
                })
                print(f"  提取2.1 外购电力（基于位置）: {total_formatted} tCO2e")

            if market_total:
                total = self.safe_float(market_total[2]) if len(market_total) > 2 else 0
                co2 = self.safe_float(market_total[3]) if len(market_total) > 3 else total

                total_formatted = f"{total:,.2f}" if total > 0 else "0.00"
                co2_formatted = f"{co2:,.2f}" if co2 > 0 else "0.00"

                scope2_items.append({
                    'number': '2.2',
                    'emission_source': '外购电力（基于市场）',
                    'total_green_house_gas_emissions': total_formatted,
                    'CO2_emissions': co2_formatted,
                    'CH4_emissions': '0.00',
                    'N2O_emissions': '0.00',
                    'HFCs_emissions': '0.00',
                    'PFCs_emissions': '0.00',
                    'SFs_emissions': '0.00',
                    'NF3_emissions': '0.00'
                })
                print(f"  提取2.2 外购电力（基于市场）: {total_formatted} tCO2e")

            result['scope2_items'] = scope2_items
            print(f"[范围二] 提取到范围二排放明细: {len(scope2_items)} 行")

        return result


__all__ = ['Scope2Reader']
