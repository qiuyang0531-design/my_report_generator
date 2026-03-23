"""
活动数据汇总表读取模块
=====================

从Excel中提取活动数据汇总表数据。
"""

from typing import Dict, List, Any

from .base import BaseReader


class ActivitySummaryReader(BaseReader):
    """活动数据汇总表读取器"""

    def extract_table1_table2_data(self) -> Dict[str, Any]:
        """
        从表1温室气体盘查表中提取表1和表2的数据

        Returns:
            包含表1和表2数据的字典
        """
        result = {
            'scope1_items': [],  # 表1：范围一直接排放源
            'scope2_3_items': [],  # 表2：范围二三间接排放源
        }

        # 查找表1温室气体盘查表
        table1_data = self.find_sheet_by_name('表1', '温室气体盘查表')

        if not table1_data:
            return result

        print(f"[表1表2] 找到表1温室气体盘查表: {table1_data.title}")

        # 维护类别变量用于前向填充（ffill）
        current_category = ""

        # 提取数据（从第5行开始）
        for row in table1_data.iter_rows(min_row=5):
            if len(row) < 7:
                continue

            # 获取各列数据
            seq = row[0].value  # 序号
            ghg_category = row[1].value  # GHG排放类别（第2列，索引为1）
            emission_source = row[2].value  # 排放源
            facility = row[3].value  # 设施
            boundary = row[4].value  # 组织边界/排放边界

            # 实现前向填充逻辑（ffill）
            if ghg_category:
                current_category = str(ghg_category).strip()

            # 跳过空行或标题行
            if not seq and not current_category:
                continue

            seq_str = str(seq).strip() if seq else ''
            ghg_str = current_category if current_category else ''
            source_str = str(emission_source).strip() if emission_source else ''
            facility_str = str(facility).strip() if facility else ''
            boundary_str = str(boundary).strip() if boundary else ''

            # 跳过标题行
            if seq_str == '序号' or ghg_str == 'GHG排放类别':
                continue

            # 表1：范围一
            if '范围一' in boundary_str:
                result['scope1_items'].append({
                    'name': ghg_str,
                    'number': seq_str,
                    'emission_source': source_str,
                    'facility': facility_str
                })

            # 表2：范围二三
            elif '范围二' in boundary_str or '范围三' in boundary_str:
                result['scope2_3_items'].append({
                    'name': ghg_str,
                    'number': seq_str,
                    'emission_source': source_str,
                    'facility': facility_str
                })

        print(f"[表1表2] scope1_items: {len(result['scope1_items'])} 行")
        print(f"[表1表2] scope2_3_items: {len(result['scope2_3_items'])} 行")

        return result


__all__ = ['ActivitySummaryReader']
