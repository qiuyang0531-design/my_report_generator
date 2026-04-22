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
            print(f"[范围二汇总] 正在从 {table_sheet.title} 提取数据...")
            
            # 查找所有包含"总排放量"的行
            summary_rows = []
            for row in table_sheet.iter_rows(values_only=True):
                a_val = row[0] if len(row) > 0 else None
                if a_val and isinstance(a_val, str) and '总排放量' in a_val:
                    summary_rows.append(row)
            
            print(f"[范围二汇总] 找到 {len(summary_rows)} 个总排放量汇总行")

            # 通常第一行是基于位置，第二行是基于市场
            if len(summary_rows) >= 1:
                loc_row = summary_rows[0]
                # Column B=Scope 1, C=Scope 2 Loc, D=Scope 3, E=Total Loc
                result['scope_1_emissions'] = self.safe_float(loc_row[1])
                result['scope_2_location_based_emissions'] = self.safe_float(loc_row[2])
                result['scope_3_emissions'] = self.safe_float(loc_row[3])
                result['scope_2_location'] = result['scope_2_location_based_emissions']
                result['total_emission_location'] = self.safe_float(loc_row[4])
                print(f"  [位置法] Scope1: {result['scope_1_emissions']}, Scope2: {result['scope_2_location']}, Scope3: {result['scope_3_emissions']}")

            if len(summary_rows) >= 2:
                mar_row = summary_rows[1]
                # Column B=Scope 1, C=Scope 2 Mkt, D=Scope 3, E=Total Mkt
                result['scope_2_market_based_emissions'] = self.safe_float(mar_row[2])
                result['scope_2_market'] = result['scope_2_market_based_emissions']
                result['total_emission_market'] = self.safe_float(mar_row[4])
                print(f"  [市场法] Scope2: {result['scope_2_market']}, Total: {result['total_emission_market']}")
            else:
                # 如果没有第二行，尝试计算
                scope_1 = result.get('scope_1_emissions') or 0
                scope_2_mkt = result.get('scope_2_market_based_emissions') or 0
                scope_3 = result.get('scope_3_emissions') or 0
                result['total_emission_market'] = scope_1 + scope_2_mkt + scope_3

        return result

    def _extract_scope2_items(self) -> Dict[str, Any]:
        """
        从温室气体盘查清册中提取范围二详细数据
        """
        result = {'scope2_items': []}

        # 查找温室气体盘查清册表
        inventory_sheet = self.find_sheet_by_name('盘查清册', '清册')

        if not inventory_sheet:
            print("[范围二详细] 未找到温室气体盘查清册表")
            return result

        print(f"[范围二详细] 找到温室气体盘查清册表: {inventory_sheet.title}")
        scope2_items = []

        # 遍历所有行，查找编号以 "2." 开头的行
        for row in inventory_sheet.iter_rows(min_row=14):
            if len(row) < 13:
                continue

            # Excel结构：A=空, B=编号/类别名, C=排放源, D=排放设施, E=备注, F=总排放量, G=CO2, H=CH4, I=N2O, J=HFCs, K=PFCs, L=SF6, M=NF3
            col_b = row[1].value        # 编号或类别名
            if not col_b:
                continue
                
            col_b_str = str(col_b).strip()
            
            # 检查是否是以 "2." 开头的编号（范围二数据）
            if col_b_str.startswith('2.'):
                item = {
                    'number': col_b_str,
                    'emission_source': self.safe_str(row[2].value),
                    'facility': self.safe_str(row[3].value),
                    'note': self.safe_str(row[4].value),
                    'total_green_house_gas_emissions': self.safe_float(row[5].value),
                    'CO2_emissions': self.safe_float(row[6].value),
                    'CH4_emissions': self.safe_float(row[7].value),
                    'N2O_emissions': self.safe_float(row[8].value),
                    'HFCs_emissions': self.safe_float(row[9].value),
                    'PFCs_emissions': self.safe_float(row[10].value),
                    'SFs_emissions': self.safe_float(row[11].value),
                    'SF6_emissions': self.safe_float(row[11].value),
                    'NF3_emissions': self.safe_float(row[12].value),
                }
                scope2_items.append(item)

        result['scope2_items'] = scope2_items
        print(f"[范围二详细] 提取到范围二排放明细: {len(scope2_items)} 行")

        return result


__all__ = ['Scope2Reader']
