"""
协议定义模块
=====================

定义所有表格类型的识别规则和字段映射。
"""

from .config import TableProtocol, FieldMapping, _PROTOCOL_ORDER


# ==================== 核心协议配置 ====================

TABLE_PROTOCOLS: Dict[str, TableProtocol] = {
    'EmissionFactorProtocol': TableProtocol(
        name='排放因子表',
        description='包含低位发热量、氧化率、基于热值排放系数等信息的表格',
        required_keywords={'低位发热量', '氧化率'},
        optional_keywords={'基于热值排放系数', '排放因子', 'GHG排放类别', '计算值', '排放系数'},
        field_mappings={
            'category': FieldMapping('GHG排放类别', '', 'str'),
            'number': FieldMapping('编号', '', 'str'),
            'emission_source': FieldMapping('排放源', '', 'str'),
            'facility': FieldMapping('设施', '', 'str'),
            'ncv': FieldMapping('低位发热量', 0, 'float'),
            'unit': FieldMapping('单位', '', 'str'),
            'ox_rate': FieldMapping('氧化率', 0, 'float'),
            # 基于热值排放系数（CO2/CH4/N2O来自"基于热值排放系数"部分）
            'CO2_emission_cv_factor': FieldMapping('CO2', 0, 'float', ['基于热值']),
            'CH4_emission_cv_factor': FieldMapping('CH4', 0, 'float', ['基于热值']),
            'N2O_emission_cv_factor': FieldMapping('N2O', 0, 'float', ['基于热值']),
            # 排放因子（CO2/CH4/N2O来自"排放因子"部分）
            'CO2_emission_factor': FieldMapping('CO2', 0, 'float', ['排放因子']),
            'CH4_emission_factor': FieldMapping('CH4', 0, 'float', ['排放因子']),
            'N2O_emission_factor': FieldMapping('N2O', 0, 'float', ['排放因子']),
        },
        output_var='pro_ef_items',
        ffill_fields=['category'],
    ),

    'GWPProtocol': TableProtocol(
        name='GWP值表',
        description='全球变暖潜势(GWP)值参考表',
        required_keywords={'GWP'},
        optional_keywords={'GWP(HFCs)', 'GWP(PFCs)', '工业名称', '中文名称', '化学分子式'},
        field_mappings={
            'gas_name': FieldMapping('工业名称', '', 'str'),
            'chinese_name': FieldMapping('中文名称/化学分子式', '', 'str'),
            'formula': FieldMapping('中文名称/化学分子式', '', 'str'),
            'composition_ratio': FieldMapping('组成比例', None, 'float'),
            'gwp_value': FieldMapping('GWP', 0, 'float'),
            'gwp_hfcs': FieldMapping('GWP(HFCs)', None, 'float'),
            'gwp_pfcs': FieldMapping('GWP(PFCs)', None, 'float'),
            'source': FieldMapping('来源', '', 'str'),
            'note': FieldMapping('备注', '', 'str'),
        },
        output_var='gwp_items',
    ),

    'GHGInventoryProtocol': TableProtocol(
        name='温室气体盘查表',
        description='温室气体排放盘查汇总表',
        required_keywords={'GHG排放类别', '排放量'},
        optional_keywords={'排放源', '设施', 'GWP', 'EF', '活动数据'},
        field_mappings={
            'category': FieldMapping('GHG排放类别', '', 'str'),
            'emission_source': FieldMapping('排放源', '', 'str'),
            'facility': FieldMapping('设施', '', 'str'),
            'activity_data': FieldMapping('活动数据', 0, 'float'),
            'activity_data_unit': FieldMapping('单位', '', 'str'),
            'emission_factor': FieldMapping('EF', 0, 'float'),
            'gwp': FieldMapping('GWP', 1, 'float'),
            'emissions': FieldMapping('排放量', 0, 'float'),
            'emissions_unit': FieldMapping('tCO2e', '', 'str'),
        },
        output_var='ghg_inventory_items',
    ),

    'Scope1EmissionsProtocol': TableProtocol(
        name='范围一直接排放源表',
        description='从附表1-温室气体盘查表中提取范围一直接排放源数据',
        required_keywords={'编号', 'GHG排放类别'},
        optional_keywords={'排放源', '设施', 'CO2', 'CH4', 'N2O', 'HFCs', 'PFCs', 'SF6', 'NF3', '总量'},
        field_mappings={
            'number': FieldMapping('编号', '', 'str'),
            'category': FieldMapping('GHG排放类别', '', 'str'),
            'emission_source': FieldMapping('排放源', '', 'str'),
            'facility': FieldMapping('设施', '', 'str'),
            'CO2_emissions': FieldMapping('CO2', 0, 'float', ['CO2排放量']),
            'CH4_emissions': FieldMapping('CH4', 0, 'float', ['CH4排放量']),
            'N2O_emissions': FieldMapping('N2O', 0, 'float', ['N2O排放量']),
            'HFCs_emissions': FieldMapping('HFCs', 0, 'float', ['HFCs排放量']),
            'PFCs_emissions': FieldMapping('PFCs', 0, 'float', ['PFCs排放量']),
            'SF6_emissions': FieldMapping('SF6', 0, 'float', ['SF6排放量']),
            'NF3_emissions': FieldMapping('NF3', 0, 'float', ['NF3排放量']),
            'total_green_house_gas_emissions': FieldMapping('总量', 0, 'float', ['总计']),
        },
        output_var='scope1_emissions_items',
        ffill_fields=['category'],
        min_match_ratio=0.5,  # 提高匹配度要求，避免与排放因子表混淆
    ),

    'ActivitySummaryProtocol': TableProtocol(
        name='活动数据汇总表',
        description='基于位置的活动数据汇总表',
        required_keywords={'编号', '排放源'},  # 修正：实际表头是"编号"不是"序号"
        optional_keywords={'GHG', '基于位置', 'CO2', 'CH4', 'N2O', '报告边界', '活动数据'},
        field_mappings={
            'number': FieldMapping('编号', '', 'str', ['序号']),  # 修正：实际表头是"编号"
            'category': FieldMapping('GHG排放类别', '', 'str'),
            'emission_source': FieldMapping('排放源', '', 'str'),
            'report_boundary': FieldMapping('报告边界', '', 'str'),
            'activity_data': FieldMapping('活动数据', 0, 'float'),  # 主表头中的"活动数据"列
            'unit': FieldMapping('计量单位', '', 'str'),  # 子表头中的"计量单位"列
            'CO2_emissions': FieldMapping('CO2', 0, 'float'),
            'CH4_emissions': FieldMapping('CH4', 0, 'float'),
            'N2O_emissions': FieldMapping('N2O', 0, 'float'),
            'HFCs_emissions': FieldMapping('HFCs', 0, 'float'),
            'PFCs_emissions': FieldMapping('PFCs', 0, 'float'),
            'SF6_emissions': FieldMapping('SF6', 0, 'float'),
            'NF3_emissions': FieldMapping('NF3', 0, 'float'),
            'total_green_house_gas_emissions': FieldMapping('总计', 0, 'float', ['总量']),
        },
        output_var='activity_summary_items',
        ffill_fields=['category'],
        min_match_ratio=0.2,  # 降低匹配度要求，因为活动数据汇总表的差异较大
        sheet_name_patterns=['活动数据汇总表', '位置'],  # 基于位置的活动数据汇总表
        header_rows_to_check=3,  # 活动数据汇总表需要检查3行表头
    ),

    'ActivitySummaryMarketProtocol': TableProtocol(
        name='活动数据汇总表（市场法）',
        description='基于市场的活动数据汇总表',
        required_keywords={'编号', '排放源'},  # 修正：实际表头是"编号"不是"序号"
        optional_keywords={'GHG', '基于市场', 'CO2', 'CH4', 'N2O', '报告边界', '活动数据'},
        field_mappings={
            'number': FieldMapping('编号', '', 'str', ['序号']),  # 修正：实际表头是"编号"
            'category': FieldMapping('GHG排放类别', '', 'str'),
            'emission_source': FieldMapping('排放源', '', 'str'),
            'report_boundary': FieldMapping('报告边界', '', 'str'),
            'activity_data': FieldMapping('活动数据', 0, 'float'),  # 主表头中的"活动数据"列
            'unit': FieldMapping('计量单位', '', 'str'),  # 子表头中的"计量单位"列
            'CO2_emissions': FieldMapping('CO2', 0, 'float'),
            'CH4_emissions': FieldMapping('CH4', 0, 'float'),
            'N2O_emissions': FieldMapping('N2O', 0, 'float'),
            'HFCs_emissions': FieldMapping('HFCs', 0, 'float'),
            'PFCs_emissions': FieldMapping('PFCs', 0, 'float'),
            'SF6_emissions': FieldMapping('SF6', 0, 'float'),
            'NF3_emissions': FieldMapping('NF3', 0, 'float'),
            'total_green_house_gas_emissions': FieldMapping('总计', 0, 'float', ['总量']),
        },
        output_var='activity_summary_market_items',
        ffill_fields=['category'],
        min_match_ratio=0.2,  # 降低匹配度要求，因为活动数据汇总表的差异较大
        sheet_name_patterns=['活动数据汇总表', '市场'],  # 基于市场的活动数据汇总表
        header_rows_to_check=3,  # 活动数据汇总表需要检查3行表头
    ),
}


__all__ = [
    'TABLE_PROTOCOLS',
    '_PROTOCOL_ORDER',
]
