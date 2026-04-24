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
    ReductionActionReader,
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
        self.reduction_action_reader = ReductionActionReader(self.workbook)

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

        # ========== 提取减排措施数据 ==========
        reduction_action_data = self.reduction_action_reader.extract()
        result.update(reduction_action_data)

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

        # ================= V18 最终整合版：含高压开关与耗氧池修正 ================= 
        
        def get_clean_desc(text): 
            if not text: return "" 
            import re 
            # 压缩多余空格并清理换行，确保 Word 缩进正常 
            return re.sub(r'\s+', ' ', str(text).strip().replace('\n', '').replace('\r', '')) 

        # 1. 基础映射与变量准备 
        all_table1 = result.get('scope1_items', []) + result.get('scope2_3_items', []) 
        source_map = {i.get('emission_source'): i.get('data_source') for i in all_table1 if i.get('emission_source')} 
        
        ef_items = result.get('pro_ef_items', []) 
        ef_ref_map = {i.get('emission_source'): i.get('emission_source_reference') for i in ef_items if i.get('emission_source')} 
        
        company = result.get('company_name', '大冶特殊钢有限公司') 
        period = result.get('reporting_period', '2025年度') 
        gwp_ref = result.get('GWP_Value_Reference_Document') or "《2021年IPCC第六次评估报告AR6》" 

        # 2. 固体燃料识别函数 
        def is_solid(name): 
            n = str(name) 
            return any(s in n for s in ['煤', '焦', '炭', '丁', '屑']) and \
                   not any(e in n for e in ['煤气', '油', '生产', '外售']) 

        # 预扫描固体燃料名单 
        active_solids = [n for n in source_map.keys() if is_solid(n)] 
        fuel_list_str = "、".join(sorted(list(set([n.replace('燃烧','').strip() for n in active_solids])))) 
        solid_ds = "、".join(list(dict.fromkeys([source_map[n] for n in active_solids if source_map[n]]))) 

        for scope in ['scope_1', 'scope_2', 'scope_3']: 
            methods = result.get('quantification_methods', {}).get(scope, {}) 
            for m_key, m_info in methods.items(): 
                # 强制名称替换：厌氧池 -> 耗氧池 
                m_name = m_info.get('name', '').replace('厌氧池', '耗氧池') 
                m_info['name'] = m_name 
                
                fuel_clean = m_name.replace('燃烧','').replace('外售','').replace('排放','').strip() 

                # 动态获取当前项的 EF 引用源 
                ref_key = next((k for k in ef_ref_map.keys() if k == m_name), 
                              next((k for k in ef_ref_map.keys() if k in m_name or m_name in k), None)) 
                curr_ef_ref = ef_ref_map.get(ref_key) or "相关核算指南" 

                # 动态获取 H 列来源 
                s_key = next((k for k in source_map.keys() if k in m_name or m_name in k), None) 
                ds = source_map.get(s_key, "相关报表") 

                # ===================================================================== 
                # 专项分支拦截 
                # ===================================================================== 

                # 1. 高压开关 (SF6 逸散) - 新增逻辑 
                if any(x in m_name for x in ["高压开关", "SF6"]): 
                    m_info['ad'] = get_clean_desc(f"来源于{company}提供{ds}{fuel_clean}填充SF6铭牌额定量的统计。") 
                    m_info['ef'] = get_clean_desc(f"参考GB/T8905-2008 六氟化硫电气设备中气体管理和检测导则9.3，逸散率取值0.5%。") 

                # 2. 废水处理 - 化粪池 (BOD类) 
                elif "化粪池" in m_name: 
                    m_info['ad'] = get_clean_desc(f"来源于{company}提供{ds}{period}员工出勤总工时推算BOD排放量的统计。") 
                    m_info['ef'] = get_clean_desc( 
                        f"EF=Bo*MCF=0.18（kgCH4/kgBOD）；所需的参数包括Bo甲烷产生最大能力、MCF甲烷修正因子和人均BOD产生量，" 
                        f"分别来源于IPCC《2006 年国家温室气体清单指南》第5卷第6章表6.2、表6.3和表6.4。其中Bo取缺省因子0.6，" 
                        f"MCF取0.3，因{company}的生活废水工业废水处理同在耗氧处理厂中，管理不完善而保守选取0.3。" 
                    ) 
                
                # 3. 废水处理 - 耗氧池 (COD类) 
                elif any(x in m_name for x in ["耗氧池", "废水处理"]): 
                    m_info['ad'] = get_clean_desc(f"来源于{company}提供{ds}{period}生产过程产生的工业废水处理量及进出口COD浓度的统计。") 
                    m_info['ef'] = get_clean_desc( 
                        f"EF=Bo*MCF=0.075（kgCH4/kgCOD）；所需的参数包括Bo甲烷产生最大能力、MCF甲烷修正因子，" 
                        f"分别来源于IPCC《2006 年国家温室气体清单指南》第5卷第6章公式6.1、表6.8。其中Bo取缺省因子0.25，" 
                        f"MCF取0.3，因{company}的生活废水工业废水处理同在耗氧处理厂中，管理不完善保守选取0.3。" 
                    ) 

                # 4. 空调/制冷剂 
                elif any(x in m_name for x in ['空调', '制冷剂', 'R32', '加氟', '逸散']): 
                    m_info['ad'] = get_clean_desc(f"来源于{company}提供{ds}{period}{fuel_clean}的统计。") 
                    m_info['ef'] = get_clean_desc(f"所需的参数为设备制冷剂充装量，量化采用质量平衡法，排放因子为 1kgGHG/kg。数据来源于{curr_ef_ref}。") 

                # --- 分支 A：固体燃料 --- 
                elif is_solid(m_name) or "固体燃料" in m_name: 
                    if not fuel_list_str: continue 
                    m_info['name'] = f"固体燃料（{fuel_list_str}燃烧）" 
                    # 应用新模板：来源于[公司]提供[来源][周期][燃料]的统计 
                    ad_text = f"来源于{company}提供{solid_ds or '相关报表'}{period}{fuel_list_str}的统计。" 
                    m_info['ad'] = get_clean_desc(ad_text) 
                    
                    m_info['ef'] = get_clean_desc( 
                        f"所需的参数包括{fuel_list_str}低位发热量、单位热值含碳量、碳氧化率；" 
                        f"数据来源于{curr_ef_ref} 附表A.1常用化石燃料相关参数缺省值；" 
                        f"固体燃料燃烧产生CO2、CH4、N2O三类温室气体热值排放系数来源于《IPCC-2006缺省值》，GWP值来源于{gwp_ref}。" 
                    ) 

                # --- 分支 B：副产品外售 --- 
                elif "副产品" in m_name or "外售" in m_name: 
                    active_bp = [k for k in source_map.keys() if "外售" in k] 
                    fuels = [k.replace('外售','').replace('燃烧','').strip() for k in active_bp if '煤气' in k] 
                    mats = [k.replace('外售','').strip() for k in active_bp if any(x in k for x in ['苯', '油'])] 
                    bp_ds = "、".join(list(dict.fromkeys([source_map[k] for k in active_bp if source_map[k]]))) 
                    bp_names = "、".join(fuels + mats) 

                    if not active_bp: continue 
                    m_info['name'] = f"副产品外售（{bp_names}）" 
                    # 应用新模板 
                    ad_text = f"来源于{company}提供{bp_ds or '相关报表'}{period}{bp_names}的统计。" 
                    m_info['ad'] = get_clean_desc(ad_text) 
                    
                    ef_segs = [] 
                    if fuels: 
                        f_ref = next((v for k,v in ef_ref_map.items() if '煤气' in k), curr_ef_ref) 
                        ef_segs.append(f"{'、'.join(fuels)}量化所需的参数包括低位发热量、碳氧化率；数据来源于{f_ref}附表A.1；" 
                                       f"{'、'.join(fuels)}燃烧产生CO2、CH4、N2O三类温室气体热值排放系数来源于《IPCC-2006缺省值》") 
                    if mats: 
                        m_ref = next((v for k,v in ef_ref_map.items() if any(x in k for x in ['苯', '油'])), curr_ef_ref) 
                        ef_segs.append(f"{'、'.join(mats)}量化所需的参数为活动数据及对应的单位排放因子，数据来源于{m_ref}") 
                    
                    m_info['ef'] = get_clean_desc("；".join(ef_segs) + f"；GWP值来源于{gwp_ref}。") 

                # --- 分支 C：常规项（天然气、汽油等） --- 
                else: 
                    # 动态匹配 AD 来源 (H列) 
                    s_key = next((k for k in source_map.keys() if k in m_name or m_name in k), None) 
                    ds = source_map.get(s_key, "相关报表") 
                    
                    # 应用新模板 
                    ad_text = f"来源于{company}提供{ds}{period}{fuel_clean}的统计。" 
                    m_info['ad'] = get_clean_desc(ad_text) 
                    
                    if any(f in m_name for f in ['气', '油', '燃']): 
                        m_info['ef'] = get_clean_desc(f"所需的参数包括{fuel_clean}低位发热量、碳氧化率；数据来源于{curr_ef_ref} 附表A.1常用化石燃料相关参数缺省值；" 
                                                     f"{fuel_clean}燃烧产生CO2、CH4、N2O三类温室气体热值排放系数来源于《IPCC-2006缺省值》，GWP值来源于{gwp_ref}。") 
                    elif '电力' in m_name: 
                        m_info['ef'] = get_clean_desc(f"所需的参数为外购电力二氧化碳排放因子；数据来源于{curr_ef_ref}；GWP值来源于{gwp_ref}。") 
                    else: 
                        m_info['ef'] = get_clean_desc(f"所需的参数为该类别排放因子；数据来源于{curr_ef_ref}；GWP值来源于{gwp_ref}。") 

        # =====================================================================

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
                mapped_item['emission_unit_dir'] = item.get('unit', '')
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
                    # 模板中这些字段可能没有后缀 (Table 36 验证)
                    mapped_item['CO2_emission_cv_factor'] = item.get('CO2_emission_cv_factor', '')
                    mapped_item['CH4_emission_cv_factor'] = item.get('CH4_emission_cv_factor', '')
                    mapped_item['N2O_emission_cv_factor'] = item.get('N2O_emission_cv_factor', '')
                    mapped_item['CO2_emission_factor'] = item.get('CO2_emission_factor', '')
                    mapped_item['CH4_emission_factor'] = item.get('CH4_emission_factor', '')
                    mapped_item['N2O_emission_factor'] = item.get('N2O_emission_factor', '')
                    
                    # 同时保留带后缀的版本，以防万一
                    mapped_item[f'CO2_emission_cv_factor_{cat_prefix}'] = item.get('CO2_emission_cv_factor', '')
                    mapped_item[f'CH4_emission_cv_factor_{cat_prefix}'] = item.get('CH4_emission_cv_factor', '')
                    mapped_item[f'N2O_emission_cv_factor_{cat_prefix}'] = item.get('N2O_emission_cv_factor', '')
                    mapped_item[f'CO2_emission_factor_{cat_prefix}'] = item.get('CO2_emission_factor', '')
                    mapped_item[f'CH4_emission_factor_{cat_prefix}'] = item.get('CH4_emission_factor', '')
                    mapped_item[f'N2O_emission_factor_{cat_prefix}'] = item.get('N2O_emission_factor', '')

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
