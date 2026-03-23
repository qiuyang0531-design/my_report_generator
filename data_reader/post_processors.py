"""
后处理模块
=====================

数据分组和转换的后处理函数。
"""

from typing import Dict, List, Any


def group_by_emission_category(items: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """
    将排放因子数据按类别分组

    分组规则：
    - 固定燃烧 -> scope1_stationary_combustion_emissions_items
    - 移动燃烧 -> scope1_mobile_combustion_emissions_items
    - 制程排放 -> scope1_process_emissions_items
    - 逸散排放 -> scope1_fugitive_emissions_items
    """
    grouped = {
        'scope1_stationary_combustion_emissions_items': [],
        'scope1_mobile_combustion_emissions_items': [],
        'scope1_fugitive_emissions_items': [],
        'scope1_process_emissions_items': [],
    }

    for item in items:
        category = item.get('category', '')
        category_normalized = category.replace(' ', '')

        if '固定燃烧' in category_normalized:
            grouped['scope1_stationary_combustion_emissions_items'].append(item)
        elif '移动燃烧' in category_normalized or '移动汽油' in category_normalized or '移动柴油' in category_normalized:
            grouped['scope1_mobile_combustion_emissions_items'].append(item)
        elif '制冷产品加工使用等排放' in category_normalized or '制程' in category_normalized:
            grouped['scope1_process_emissions_items'].append(item)
        elif '逸散' in category_normalized:
            grouped['scope1_fugitive_emissions_items'].append(item)

    return grouped


def group_scope1_emissions(items: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """
    将范围一排放数据按类别分组

    与 group_by_emission_category 相同的分组逻辑
    """
    return group_by_emission_category(items)


__all__ = [
    'group_by_emission_category',
    'group_scope1_emissions',
]
