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
        result = {'scope1_items': [], 'scope2_3_items': []} 
        table1_data = self.find_sheet_by_name('表1', '温室气体盘查表') 
        if not table1_data: return result 

        current_category = "" 
        # 提取数据（从第5行开始） 
        for row in table1_data.iter_rows(min_row=5): 
            # 确保至少有8列（包含H列，索引为7） 
            if len(row) < 8: continue 

            ghg_category = row[1].value 
            if ghg_category: current_category = str(ghg_category).strip() 

            boundary_str = str(row[4].value).strip() if row[4].value else '' 
            
            # 封装数据项，增加 data_source (H列) 
            item = { 
                'name': current_category, 
                'number': str(row[0].value).strip() if row[0].value else '', 
                'emission_source': str(row[2].value).strip() if row[2].value else '', 
                'facility': str(row[3].value).strip() if row[3].value else '', 
                'data_source': str(row[7].value).strip() if row[7].value else '' # 核心修改 
            } 

            if '范围一' in boundary_str: 
                result['scope1_items'].append(item) 
            elif '范围二' in boundary_str or '范围三' in boundary_str: 
                result['scope2_3_items'].append(item) 
        
        print(f"[表1表2] scope1_items: {len(result['scope1_items'])} 行")
        print(f"[表1表2] scope2_3_items: {len(result['scope2_3_items'])} 行")

        return result


__all__ = ['ActivitySummaryReader']
