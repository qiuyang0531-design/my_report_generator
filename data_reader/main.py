"""
主入口模块
=====================

ExcelDataReaderRefactored - 重构后的Excel数据读取器主类

这是唯一的高层接口，自动识别所有表格并提取数据。
"""

import openpyxl
from typing import Dict, Any

# 尝试导入 ReportConfig 以支持 quantification_methods
try:
    from report_config import ReportConfig
    HAS_REPORT_CONFIG = True
except ImportError:
    HAS_REPORT_CONFIG = False

from .protocols import TABLE_PROTOCOLS
from .fingerprint import TableFingerprint
from .extractor import ProtocolExtractor
from .post_processors import group_by_emission_category, group_scope1_emissions
from .readers import (
    BaseReader,
    BasicInfoReader,
    Scope1Reader,
    Scope2Reader,
    Scope3Reader,
    EmissionFactorReader,
    ActivitySummaryReader,
)


class ExcelDataReaderRefactored(BaseReader):
    """
    重构后的Excel数据读取器

    这是唯一的高层接口，自动识别所有表格并提取数据
    """

    def __init__(self, file_path: str):
        """
        初始化数据读取器

        Args:
            file_path: Excel文件路径
        """
        self.file_path = file_path
        self.workbook = openpyxl.load_workbook(file_path, data_only=True)
        self.extractor = ProtocolExtractor()
        self.fingerprint = TableFingerprint()

        # 初始化专项读取器
        self.basic_info_reader = BasicInfoReader(self.workbook)
        self.scope1_reader = Scope1Reader(self.workbook)
        self.scope2_reader = Scope2Reader(self.workbook)
        self.scope3_reader = Scope3Reader(self.workbook)
        self.emission_factor_reader = EmissionFactorReader(self.workbook)
        self.activity_summary_reader = ActivitySummaryReader(self.workbook)

    def get_all_context(self) -> Dict[str, Any]:
        """
        获取所有渲染上下文数据

        这是唯一的高层接口，自动识别所有表格并提取数据

        Returns:
            包含所有提取数据的字典
        """
        result = {}

        print(f"[数据读取] 开始处理工作簿，共 {len(self.workbook.sheetnames)} 个工作表")

        # ========== 首先提取基本信息 ==========
        basic_info = self.basic_info_reader.extract()
        result.update(basic_info)

        # ========== 提取范围三类别数据 ==========
        scope3_data = self.scope3_reader.extract_all()
        result.update(scope3_data)

        # ========== 提取范围二数据 ==========
        scope2_data = self.scope2_reader.extract_all()
        result.update(scope2_data)

        # ========== 添加 quantification_methods ==========
        if HAS_REPORT_CONFIG and result.get('company_name'):
            report_config = ReportConfig(
                company_name=result['company_name'],
                reporting_period=result.get('reporting_period', '2024年')
            )
            result['quantification_methods'] = report_config.get_quantification_methods()
        else:
            result['quantification_methods'] = {}

        # ========== 提取范围三详细数据（从温室气体盘查清册）==========
        scope3_detail_data = self._extract_scope3_detail_data()
        result.update(scope3_detail_data)

        # ========== 提取表1和表2的数据（从表1温室气体盘查表）==========
        table1_table2_data = self.activity_summary_reader.extract_table1_table2_data()
        result.update(table1_table2_data)

        # ========== 提取范围一详细数据（从温室气体盘查清册表）==========
        scope1_detail_data = self.scope1_reader.extract_all()
        result.update(scope1_detail_data)

        # ========== 遍历所有工作表，识别并提取表格数据 ==========
        for sheet in self.workbook.worksheets:
            sheet_name = sheet.title

            # 识别表格类型
            protocol_name = self.fingerprint.identify(sheet, sheet_name)

            if protocol_name:
                protocol = TABLE_PROTOCOLS[protocol_name]
                output_var = protocol.output_var

                # 排放因子表特殊处理：多子表提取
                if protocol_name == 'EmissionFactorProtocol' and '附表2-EF' in sheet_name:
                    data_items = self.emission_factor_reader.extract_all()
                else:
                    # 提取数据
                    data_items = self.extractor.extract_from_sheet(sheet, protocol_name)

                # 存储结果
                result[output_var] = data_items
                print(f"[数据读取] {sheet_name} -> {output_var}: {len(data_items)} 行")

        # 确保所有输出变量都被初始化
        for protocol in TABLE_PROTOCOLS.values():
            if protocol.output_var not in result:
                result[protocol.output_var] = []

        # ========== 后处理：类别分组 ==========
        result = self._post_process_emission_factors(result)
        result = self._post_process_scope1_emissions(result)
        result = self._post_process_activity_summaries(result)
        result = self._post_process_scope3_ef_items(result)

        # ========== 后处理：更新 Flags 标记 ==========
        result = self._update_flags(result)

        return result

    def _post_process_emission_factors(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """后处理排放因子表数据"""
        if 'pro_ef_items' not in result or not result['pro_ef_items']:
            return result

        # 按子表类型分组数据
        combustion_items = []
        process_items = []
        fugitive_items = []

        for item in result['pro_ef_items']:
            category = item.get('category', '')

            if '固定燃烧' in category or '移动燃烧' in category:
                mapped_item = item.copy()
                mapped_item['emission_source_type_dir'] = item.get('category', '')
                mapped_item['emission_source_dir'] = item.get('emission_source', '')
                mapped_item['emission_facilities_dir'] = item.get('facility', '')
                mapped_item['ncv_dir'] = item.get('ncv', '')
                mapped_item['emission_unit_dir'] = item.get('unit', '')
                mapped_item['emission_oa_dir'] = item.get('ox_rate', '')
                combustion_items.append(mapped_item)

            elif '制程排放' in category:
                mapped_item = item.copy()
                mapped_item['emission_source_type_dir'] = item.get('category', '')
                mapped_item['emission_source_dir'] = item.get('emission_source', '')
                mapped_item['emission_facilities_dir'] = item.get('facility', '')
                process_items.append(mapped_item)

            elif '逸散排放' in category:
                mapped_item = item.copy()
                mapped_item['emission_source_type_dir'] = item.get('category', '')
                mapped_item['emission_source_dir'] = item.get('emission_source', '')
                mapped_item['emission_facilities_dir'] = item.get('facility', '')
                mapped_item['HFCs_PCFs_emission_factor'] = item.get('CO2_emission_factor', '')
                mapped_item['emission_factor'] = item.get('CO2_emission_factor', '')
                mapped_item['emission_unit_dir'] = item.get('unit', '')
                fugitive_items.append(mapped_item)

        # 设置三个表格的数据
        result['emission_factor_combustion_items'] = combustion_items
        result['emission_factor_process_items'] = process_items
        result['emission_factor_fugitive_items'] = fugitive_items
        result['emission_factor_items'] = combustion_items + process_items + fugitive_items

        print("[后处理] 排放因子表数据分组:")
        print(f"  emission_factor_combustion_items (表格22-燃烧): {len(combustion_items)} 条")
        print(f"  emission_factor_process_items (表格23-制程): {len(process_items)} 条")
        print(f"  emission_factor_fugitive_items (表格24-逸散): {len(fugitive_items)} 条")
        print(f"  emission_factor_items (总计): {len(result['emission_factor_items'])} 条")

        # 处理外购能源间接排放因子（范围二排放因子）
        scope2_ef_raw_items = [
            item for item in result.get('pro_ef_items', [])
            if '范围二' in item.get('category', '') and '外购能源' in item.get('category', '')
        ]

        indir_ef_items = []
        for item in scope2_ef_raw_items:
            mapped_item = {
                'number': item.get('number', ''),
                'emission_source_type_indir': item.get('category', ''),
                'emission_source_indir': item.get('emission_source', ''),
                'emission_facilities_indir': item.get('facility', ''),
                'elec_emission_factor': item.get('CO2_emission_factor', ''),
                'elec_emission_unit': item.get('unit', ''),
                'elec_emission_source': item.get('emission_source_reference', ''),
            }
            indir_ef_items.append(mapped_item)

        result['indir_ef_items'] = indir_ef_items
        print(f"  indir_ef_items (外购能源间接排放因子): {len(indir_ef_items)} 条")

        return result

    def _post_process_scope1_emissions(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """后处理范围一排放数据"""
        if 'scope1_emissions_items' in result and result['scope1_emissions_items']:
            grouped_data = group_scope1_emissions(result['scope1_emissions_items'])
            result.update(grouped_data)
            print("[后处理] 范围一排放按类别分组:")
            for group_name, items in grouped_data.items():
                if items:
                    print(f"  {group_name}: {len(items)} 条")

        return result

    def _post_process_activity_summaries(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """后处理活动数据汇总表"""
        # 处理基于位置的活动数据汇总表
        if 'activity_summary_items' in result:
            location_items = []
            for item in result['activity_summary_items']:
                mapped_item = {
                    'number': item.get('number', ''),
                    'emission_source_type_loc': item.get('category', ''),
                    'emission_source_type_location_based': item.get('category', ''),
                    'emission_source_loc': item.get('emission_source', ''),
                    'emission_source_location_based': item.get('emission_source', ''),
                    'report_boundary_loc': item.get('report_boundary', ''),
                    'report_boundary_location_based': item.get('report_boundary', ''),
                    'act_summary_loc': item.get('activity_data', ''),
                    'activity_data_location_based': item.get('activity_data', ''),
                    'act_summary_loc_unit': item.get('unit', ''),
                    'activity_data_unit_location_based': item.get('unit', ''),
                }
                for field_name in ['CO2_emissions', 'CH4_emissions', 'N2O_emissions',
                                   'HFCs_emissions', 'PFCs_emissions', 'SF6_emissions',
                                   'NF3_emissions', 'total_green_house_gas_emissions']:
                    if field_name in item:
                        mapped_item[field_name] = item[field_name]
                location_items.append(mapped_item)

            result['act_summary_loc'] = location_items
            result['activity_summary_location_items'] = location_items
            print(f"[后处理] act_summary_loc (基于位置): {len(location_items)} 行")

        # 处理基于市场的活动数据汇总表
        if 'activity_summary_market_items' in result:
            market_items = []
            for item in result['activity_summary_market_items']:
                mapped_item = {
                    'number': item.get('number', ''),
                    'emission_source_type_mar': item.get('category', ''),
                    'emission_source_type_market_based': item.get('category', ''),
                    'emission_source_mar': item.get('emission_source', ''),
                    'emission_source_market_based': item.get('emission_source', ''),
                    'report_boundary_mar': item.get('report_boundary', ''),
                    'report_boundary_market_based': item.get('report_boundary', ''),
                    'act_summary_mar': item.get('activity_data', ''),
                    'activity_data_market_based': item.get('activity_data', ''),
                    'act_summary_mar_unit': item.get('unit', ''),
                    'activity_data_unit_market_based': item.get('unit', ''),
                }
                for field_name in ['CO2_emissions', 'CH4_emissions', 'N2O_emissions',
                                   'HFCs_emissions', 'PFCs_emissions', 'SF6_emissions',
                                   'NF3_emissions', 'total_green_house_gas_emissions']:
                    if field_name in item:
                        mapped_item[field_name] = item[field_name]
                market_items.append(mapped_item)

            result['act_summary_mar'] = market_items
            print(f"[后处理] act_summary_mar (基于市场): {len(market_items)} 行")

        return result

    def _post_process_scope3_ef_items(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """后处理范围三所有类别排放因子（cat1-cat15）"""
        if 'pro_ef_items' not in result:
            return result

        # 按类别分组，提取类别编号
        category_groups = {}
        for item in result.get('pro_ef_items', []):
            category = item.get('category', '')
            cat_num = None

            # 尝试从类别名称中提取编号 (新格式)
            for cat_id in range(15, 0, -1):
                if f'范围三 类别{cat_id}' in category or f'范围三类别{cat_id}' in category or f'范围3 类别{cat_id}' in category:
                    cat_num = cat_id
                    break

            if not cat_num:
                # 使用旧格式的映射 (向后兼容)
                legacy_mapping = {
                    '外购商品和服务的上游排放': 1,
                    '资本货物': 2,
                    '范围一、二之外燃料和能源相关的活动产生的排放': 3,
                    '上下游运输配送产生的排放': 4,
                    '运营中产生的废物排放': 5,
                    '商务旅行产生的排放': 6,
                    '员工通勤': 7,
                    '上游租赁资产': 8,
                    '下游运输配送': 9,
                    '外销产品加工': 10,
                    '外销产品使用': 11,
                    '外售产品报废': 12,
                }
                cat_num = legacy_mapping.get(category)

            if cat_num and 1 <= cat_num <= 15:
                if cat_num not in category_groups:
                    category_groups[cat_num] = []
                category_groups[cat_num].append(item)

        # 为所有类别1-15创建对应的变量
        for cat_num in range(1, 16):
            cat_prefix = f'cat{cat_num}'
            items = category_groups.get(cat_num, [])

            cat_ef_items = []
            for item in items:
                mapped_item = {
                    'number': item.get('number', ''),
                    f'emission_source_type_{cat_prefix}': item.get('category', ''),
                    f'emission_source_{cat_prefix}': item.get('emission_source', ''),
                    f'emission_name_{cat_prefix}': item.get('activity_name', ''),
                    f'emission_geo_{cat_prefix}': item.get('geography', ''),
                    f'{cat_prefix}_emission_factor': item.get('CO2_emission_factor', ''),
                    f'{cat_prefix}_emission_unit': item.get('unit', ''),
                    f'{cat_prefix}_emission_source': item.get('emission_source_reference', ''),
                }

                # 检查是否是燃烧表格式
                if item.get('ncv') is not None and item.get('ncv') != 0:
                    mapped_item[f'ncv_{cat_prefix}'] = item.get('ncv', '')
                    mapped_item[f'emission_unit_{cat_prefix}'] = item.get('unit', '')
                    mapped_item[f'emission_oa_{cat_prefix}'] = item.get('ox_rate', '')
                    mapped_item[f'CO2_emission_cv_factor'] = item.get('CO2_emission_cv_factor', '')
                    mapped_item[f'CH4_emission_cv_factor'] = item.get('CH4_emission_cv_factor', '')
                    mapped_item[f'N2O_emission_cv_factor'] = item.get('N2O_emission_cv_factor', '')
                    mapped_item[f'CO2_emission_factor'] = item.get('CO2_emission_factor', '')
                    mapped_item[f'CH4_emission_factor'] = item.get('CH4_emission_factor', '')
                    mapped_item[f'N2O_emission_factor'] = item.get('N2O_emission_factor', '')

                cat_ef_items.append(mapped_item)

            result[f'{cat_prefix}_ef_items'] = cat_ef_items
            if items:
                category_name = items[0].get('category', '')[:40]
                print(f"  {cat_prefix}_ef_items (类别{cat_num}): {len(cat_ef_items)} 条")
            else:
                print(f"  {cat_prefix}_ef_items (类别{cat_num}): 0 条 - 无数据")

        return result

    def _extract_scope3_detail_data(self) -> Dict[str, Any]:
        """从温室气体盘查清册中提取范围三详细数据"""
        # 这个方法在 Scope3Reader 中已实现
        # 这里保留空实现，避免重复代码
        return {}

    def _update_flags(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """更新 Flags 标记系统"""
        if 'flags' not in data:
            data['flags'] = {}

        def safe_float(value):
            try:
                return float(value)
            except (ValueError, TypeError):
                return 0.0

        data['flags']['has_scope_1'] = safe_float(data.get('scope_1_emissions', 0)) > 0
        data['flags']['has_scope_2_location'] = safe_float(data.get('scope_2_location_based_emissions', 0)) > 0
        data['flags']['has_scope_2_market'] = safe_float(data.get('scope_2_market_based_emissions', 0)) > 0
        data['flags']['has_scope_3'] = safe_float(data.get('scope_3_emissions', 0)) > 0

        for i in range(1, 16):
            key = f'scope_3_category_{i}_emissions'
            flag_key = f'has_scope_3_category_{i}'
            data['flags'][flag_key] = safe_float(data.get(key, 0)) > 0

        return data

    def close(self):
        """关闭工作簿"""
        if self.workbook:
            self.workbook.close()


__all__ = ['ExcelDataReaderRefactored']
