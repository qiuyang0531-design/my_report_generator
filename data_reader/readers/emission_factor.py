"""
排放因子表读取模块
=====================

从Excel中提取排放因子表数据（包含多个子表）。
"""

import re
import openpyxl
from typing import Dict, List, Any

from .base import BaseReader


class EmissionFactorReader(BaseReader):
    """排放因子表读取器"""

    def extract_all(self) -> List[Dict[str, Any]]:
        """
        提取排放因子表的所有子表数据

        Returns:
            包含所有子表数据的字典列表
        """
        # 查找附表2-EF工作表
        sheet = self.find_sheet_by_name('附表2-EF')

        if not sheet:
            print("[排放因子表] 未找到附表2-EF工作表")
            return []

        return self._extract_subtables(sheet)

    def _extract_subtables(self, sheet: openpyxl.worksheet.Worksheet) -> List[Dict[str, Any]]:
        """
        从排放因子表中提取所有子表的数据

        附表2-EF 包含多个子表，每个子表以"编号"开始
        每个子表对应不同的排放类别和不同的列结构

        Returns:
            包含所有子表数据的字典列表
        """
        print(f"[排放因子表] 开始识别子表...")

        # 第一步：找到所有"编号"出现的行（子表开始位置）
        subtable_starts = []
        for row_idx in range(1, sheet.max_row + 1):
            for cell in sheet[row_idx]:
                if cell.value and str(cell.value).strip() == '编号':
                    subtable_starts.append(row_idx)
                    break

        print(f"[排放因子表] 找到 {len(subtable_starts)} 个子表")

        all_data = []

        # 第二步：处理每个子表
        for i, start_row in enumerate(subtable_starts):
            end_row = subtable_starts[i + 1] if i + 1 < len(subtable_starts) else sheet.max_row + 1

            # 读取子表的前几行来识别表头结构
            subtable_data = self._extract_single_subtable(sheet, start_row, end_row)

            if subtable_data:
                all_data.extend(subtable_data)
                print(f"[排放因子表] 子表 {i + 1}: 提取到 {len(subtable_data)} 行数据")

        print(f"[排放因子表] 总共提取到 {len(all_data)} 行数据")
        return all_data

    def _extract_single_subtable(self, sheet: openpyxl.worksheet.Worksheet,
                                 start_row: int, end_row: int) -> List[Dict[str, Any]]:
        """
        从单个排放因子子表中提取数据

        Args:
            sheet: 工作表对象
            start_row: 子表开始行
            end_row: 子表结束行

        Returns:
            该子表的数据列表
        """
        # 读取前5行来识别表头结构
        header_rows = []
        for row_idx in range(start_row, min(start_row + 5, end_row)):
            row = sheet[row_idx]
            row_data = []
            for cell in row:
                value = cell.value if cell.value is not None else ''
                row_data.append(str(value).strip())
            header_rows.append(row_data)

        # 查找第一行数据以获取类别信息
        first_data_row = None
        for row_idx in range(start_row + 2, min(start_row + 10, end_row)):
            row = sheet[row_idx]
            # 检查Col2是否有数字编号
            if len(row) > 1:
                col2_value = row[1].value if row[1].value is not None else ''
                try:
                    float(col2_value)  # 如果是数字，说明这是数据行
                    first_data_row = row
                    break
                except (ValueError, TypeError):
                    continue

        # 如果找到数据行，将其信息也添加到header_rows用于识别
        if first_data_row:
            data_row_data = []
            for cell in first_data_row[:8]:
                value = cell.value if cell.value is not None else ''
                data_row_data.append(str(value).strip())
            header_rows.append(data_row_data)

        # 识别子表类型（通过分析表头结构和数据行）
        subtable_type = self._identify_subtable_type(header_rows)

        # 根据子表类型提取数据
        return self._extract_subtable_data_by_type(sheet, start_row, end_row, subtable_type)

    def _identify_subtable_type(self, header_rows: List[List[str]]) -> str:
        """
        根据表头行识别子表类型

        支持显式类别标识：
        - "范围一 XXX" -> scope1_combustion/scope1_process/scope1_fugitive
        - "范围二 XXX" -> scope2
        - "范围三 类别N XXX" -> scope3_catN (N为1-15)

        Returns:
            子表类型
        """
        # 合并所有表头行进行分析
        all_text = ' '.join([' '.join(row) for row in header_rows])

        # 优先级1：识别范围三 "范围三 类别N" (简单字符串匹配)
        for cat_num in range(15, 0, -1):
            if f'范围三 类别{cat_num}' in all_text or f'范围三类别{cat_num}' in all_text or f'范围3 类别{cat_num}' in all_text:
                return f'scope3_cat{cat_num}'

        # 优先级2：识别范围二 "范围二"
        if '范围二' in all_text or '范围2' in all_text:
            return 'scope2'

        # 优先级3：识别范围一排放因子类型
        if '范围一' in all_text or '范围1' in all_text:
            if '低位发热量' in all_text and '氧化率' in all_text:
                return 'combustion'
            elif '制程排放' in all_text or '工艺排放' in all_text:
                return 'process'
            elif '逸散排放' in all_text or 'HFCs/PFCs' in all_text:
                return 'fugitive'
            return 'scope1_combustion'

        # 优先级4：旧格式关键词匹配（向后兼容）
        if '低位发热量' in all_text and '氧化率' in all_text:
            return 'combustion'
        elif 'CO2排放因子' in all_text and '制程排放' in all_text:
            return 'process'
        elif 'HFCs/PFCs' in all_text or 'MCF' in all_text or 'Bo' in all_text:
            return 'fugitive'
        elif '外购能源间接排放' in all_text:
            return 'scope2'

        # 优先级5：范围三类别关键词匹配（旧格式）
        scope3_keywords = {
            '外购商品和服务': 'scope3_cat1',
            '铁矿石': 'scope3_cat1',
            '资本货物': 'scope3_cat2',
            '燃料和能源相关': 'scope3_cat3',
            '上下游运输配送': 'scope3_cat4',
            '运营中产生的废物': 'scope3_cat5',
            '商务旅行': 'scope3_cat6',
            '员工通勤': 'scope3_cat7',
            '上游租赁资产': 'scope3_cat8',
            '下游运输配送': 'scope3_cat9',
            '外销产品加工': 'scope3_cat10',
            '外销产品使用': 'scope3_cat11',
            '外售产品报废': 'scope3_cat12',
        }

        for keyword, subtable_type in scope3_keywords.items():
            if keyword in all_text:
                return subtable_type

        return 'unknown'

    def _extract_subtable_data_by_type(self, sheet: openpyxl.worksheet.Worksheet,
                                       start_row: int, end_row: int,
                                       subtable_type: str) -> List[Dict[str, Any]]:
        """
        根据子表类型提取数据

        Args:
            sheet: 工作表对象
            start_row: 子表开始行
            end_row: 子表结束行
            subtable_type: 子表类型

        Returns:
            该子表的数据列表
        """
        data_items = []

        # 检查子表是否有燃烧表格式
        has_combustion_format = False
        for row_idx in range(start_row, min(start_row + 5, end_row)):
            row = sheet[row_idx]
            for cell in row:
                if cell.value and isinstance(cell.value, str) and '低位发热量' in cell.value:
                    has_combustion_format = True
                    break
            if has_combustion_format:
                break

        # 动态确定数据开始行
        data_start_row = start_row + 1
        for row_idx in range(start_row + 1, min(start_row + 6, end_row)):
            row = sheet[row_idx]
            if len(row) > 1:
                col2_value = row[1].value if row[1].value is not None else ''
                try:
                    float(col2_value)
                    data_start_row = row_idx
                    break
                except (ValueError, TypeError):
                    continue

        for row_idx in range(data_start_row, end_row):
            row = sheet[row_idx]

            # 检查是否是空行
            if not any(cell.value is not None for cell in row):
                continue

            # 检查第一列(编号)是否是数字，或者包含特定的前缀。如果不是，可能是标题或备注，跳过。
            # 这可以防止"类别11"这种标题被当作数据行提取到"类别9"中。
            number_val = row[1].value # 编号在第二列
            if number_val:
                number_str = str(number_val).strip()
                # 如果是"类别X"格式，跳过
                if number_str.startswith('类别') or '范围' in number_str:
                    continue
                # 如果不是以数字开头，且不是空，也可能是标题，跳过
                if number_str and not number_str[0].isdigit():
                    continue

            # 根据子表类型提取数据
            if subtable_type == 'combustion':
                item = self._extract_combustion_row(row)
            elif subtable_type == 'process':
                item = self._extract_process_row(row)
            elif subtable_type == 'fugitive':
                item = self._extract_fugitive_row(row)
            elif subtable_type == 'scope2':
                item = self._extract_scope2_row(row)
            elif subtable_type.startswith('scope3_cat'):
                if has_combustion_format:
                    item = self._extract_combustion_row(row)
                else:
                    item = self._extract_scope3_row(row, subtable_type)
            elif subtable_type in ['scope3_general', 'scope3_capital', 'scope3_fuel',
                                   'scope3_transport', 'scope3_waste', 'scope3_business',
                                   'scope3_commuting', 'scope3_processing', 'scope3_disposal']:
                item = self._extract_scope3_row(row, subtable_type)
            else:
                continue

            if item and item.get('category'):
                data_items.append(item)

        return data_items

    def _extract_combustion_row(self, row) -> Dict[str, Any]:
        """提取燃烧排放因子行数据"""
        try:
            def looks_like_energy_unit(value) -> bool:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return False
                s_compact = s.lower().replace(' ', '')
                if '/' not in s_compact:
                    return False
                return any(k in s_compact for k in ['gj', 'mj', 'kwh', 'nm3', 'm3', 't', 'kg'])

            unit = self.safe_get_cell(row, 6)
            if not str(unit).strip():
                for col_idx in [5, 7, 8, 4, 9, 10]:
                    candidate = self.safe_get_cell(row, col_idx)
                    if looks_like_energy_unit(candidate):
                        unit = candidate
                        break

            return {
                'number': self.safe_get_cell(row, 1),
                'category': self.safe_get_cell(row, 2),
                'emission_source': self.safe_get_cell(row, 3),
                'facility': self.safe_get_cell(row, 4),
                'ncv': self.safe_float(self.safe_get_cell(row, 5)),
                'unit': unit,
                'ox_rate': self.safe_float(self.safe_get_cell(row, 7)),
                'CO2_emission_cv_factor': self.safe_float(self.safe_get_cell(row, 8)),
                'CH4_emission_cv_factor': self.safe_float(self.safe_get_cell(row, 9)),
                'N2O_emission_cv_factor': self.safe_float(self.safe_get_cell(row, 10)),
                'CO2_emission_factor': self.safe_float(self.safe_get_cell(row, 11)),
                'CH4_emission_factor': self.safe_float(self.safe_get_cell(row, 12)),
                'N2O_emission_factor': self.safe_float(self.safe_get_cell(row, 13)),
            }
        except Exception:
            return {}

    def _extract_process_row(self, row) -> Dict[str, Any]:
        """提取制程排放因子行数据"""
        try:
            def looks_like_emission_unit(value) -> bool:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return False
                s_compact = s.lower().replace(' ', '')
                return ('co2' in s_compact) and ('/' in s_compact)

            unit = self.safe_get_cell(row, 6)
            if not str(unit).strip():
                for col_idx in [5, 7, 8, 4, 9, 10]:
                    candidate = self.safe_get_cell(row, col_idx)
                    if looks_like_emission_unit(candidate):
                        unit = candidate
                        break

            return {
                'number': self.safe_get_cell(row, 1),
                'category': self.safe_get_cell(row, 2),
                'emission_source': self.safe_get_cell(row, 3),
                'facility': self.safe_get_cell(row, 4),
                'ncv': 0,
                'unit': unit,
                'ox_rate': 0,
                'CO2_emission_cv_factor': 0,
                'CH4_emission_cv_factor': 0,
                'N2O_emission_cv_factor': 0,
                'CO2_emission_factor': self.safe_float(self.safe_get_cell(row, 5)),
                'CH4_emission_factor': 0,
                'N2O_emission_factor': 0,
            }
        except Exception:
            return {}

    def _extract_fugitive_row(self, row) -> Dict[str, Any]:
        """提取逸散排放因子行数据"""
        try:
            category = self.safe_get_cell(row, 2)
            hfcs_pfcs = self.safe_float(self.safe_get_cell(row, 5))
            unit1 = self.safe_get_cell(row, 6)
            mcf = self.safe_float(self.safe_get_cell(row, 7))
            bo = self.safe_float(self.safe_get_cell(row, 8))
            ef_value = self.safe_float(self.safe_get_cell(row, 9))
            unit2 = self.safe_get_cell(row, 10)
            source = self.safe_get_cell(row, 11)

            # 如果CH4逸散，需要使用MCF和Bo计算
            if 'CH4逸散' in category and ef_value == 0:
                if mcf > 0 and bo > 0:
                    ef_value = mcf * bo

            return {
                'number': self.safe_get_cell(row, 1),
                'category': category,
                'emission_source': self.safe_get_cell(row, 3),
                'facility': self.safe_get_cell(row, 4),
                'ncv': 0,
                'unit': unit2 or unit1,
                'ox_rate': 0,
                'CO2_emission_cv_factor': 0,
                'CH4_emission_cv_factor': 0,
                'N2O_emission_cv_factor': 0,
                'CO2_emission_factor': ef_value,
                'CH4_emission_factor': 0,
                'N2O_emission_factor': 0,
                'HFCs_PCFs_emission_factor': hfcs_pfcs,
                'MCF': mcf,
                'Bo': bo,
                'emission_factor': ef_value,
                'emission_source_dir': '',
                'emission_unit_dir': unit2 or unit1,
            }
        except Exception:
            return {}

    def _extract_scope2_row(self, row) -> Dict[str, Any]:
        """提取外购能源排放因子行数据"""
        try:
            def looks_like_emission_unit(value) -> bool:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return False
                s_compact = s.lower().replace(' ', '')
                return ('co2' in s_compact) and ('/' in s_compact)

            unit = self.safe_get_cell(row, 6)
            if not str(unit).strip():
                for col_idx in [5, 7, 8, 4, 9, 10]:
                    candidate = self.safe_get_cell(row, col_idx)
                    if looks_like_emission_unit(candidate):
                        unit = candidate
                        break

            return {
                'number': self.safe_get_cell(row, 1),
                'category': self.safe_get_cell(row, 2),
                'emission_source': self.safe_get_cell(row, 3),
                'facility': self.safe_get_cell(row, 4),
                'ncv': 0,
                'unit': unit,
                'ox_rate': 0,
                'CO2_emission_cv_factor': 0,
                'CH4_emission_cv_factor': 0,
                'N2O_emission_cv_factor': 0,
                'CO2_emission_factor': self.safe_float(self.safe_get_cell(row, 5)),
                'CH4_emission_factor': 0,
                'N2O_emission_factor': 0,
                'emission_source_reference': self.safe_get_cell(row, 7),
            }
        except Exception:
            return {}

    def _extract_scope3_row(self, row, subtable_type: str) -> Dict[str, Any]:
        """提取范围三排放因子行数据"""
        try:
            def looks_like_unit(value) -> bool:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return False
                s_lower = s.lower()
                s_compact = s_lower.replace(' ', '')
                if any(k in s_lower for k in ['supplychain', 'ecoinvent', 'ipcc', 'cpcd']):
                    return False
                return (
                    (('co2e' in s_compact) or ('co2' in s_compact))
                    and (
                        'purchaser price' in s_lower
                        or 'usd' in s_lower
                        or '/202' in s_lower
                        or 'tco2e/' in s_compact
                        or 'kgco2e/' in s_compact
                        or 'tco2/' in s_compact
                        or 'kgco2/' in s_compact
                        or ('/' in s_compact)
                    )
                )

            def looks_like_reference(value) -> bool:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return False
                if '\n' in s or '\r' in s:
                    return False
                if len(s) > 160:
                    return False
                if any(k in s for k in ['理由', '原因', '选用', '不用', '估计包含', '更具代表性']):
                    return False
                s_lower = s.lower()
                return any(
                    k in s_lower
                    for k in [
                        'supplychain',
                        'supplychainghgemissionfactors',
                        'ecoinvent',
                        'ipcc',
                        'cpcd',
                        'ghg protocol',
                        'naics',
                        'ar6',
                        'ar5',
                    ]
                )

            def extract_ecoinvent_reference(value) -> str:
                s = str(value).strip() if value is not None else ''
                if not s:
                    return ''
                m = re.search(r'(ecoinvent\s*3\.\d+\s*-\s*cut\s*-?\s*off)', s, flags=re.IGNORECASE)
                if m:
                    ref = m.group(1)
                    ref = re.sub(r'\s+', ' ', ref).strip()
                    ref = re.sub(r'\s*-\s*', '-', ref)
                    ref = ref.replace('cut-off', 'cut off')
                    return ref
                if 'ecoinvent' in s.lower():
                    return 'ecoinvent 3.10-cut off'
                return ''

            number = self.safe_get_cell(row, 1)
            raw_category = self.safe_get_cell(row, 2)
            emission_source = self.safe_get_cell(row, 3)
            activity_name = self.safe_get_cell(row, 4)

            # 根据subtable_type确定正确的类别名称
            category = raw_category
            if subtable_type.startswith('scope3_cat'):
                match = re.search(r'scope3_cat(\d+)', subtable_type)
                if match:
                    cat_num = match.group(1)
                    category = f'范围三 类别{cat_num}'
                    raw_category_str = str(raw_category) if raw_category is not None else ''
                    if raw_category_str and raw_category_str.strip() and raw_category_str != category:
                        if not raw_category_str.replace('.', '').isdigit():
                            if not emission_source or not emission_source.strip():
                                emission_source = raw_category_str

            # 确定列结构
            col5_value = self.safe_get_cell(row, 5)
            col5_str = str(col5_value) if col5_value is not None and col5_value != '' else ''
            has_geography = False
            if col5_str and col5_str.strip():
                if col5_str.startswith('='):
                    has_geography = False
                else:
                    try:
                        float(col5_value)
                        has_geography = False
                    except (ValueError, TypeError):
                        has_geography = True
            else:
                has_geography = False

            if not has_geography:
                cells = {i: self.safe_get_cell(row, i) for i in range(5, 11)}

                ef_value = 0.0
                ef_idx = None
                for idx in [5, 6, 7]:
                    try:
                        v = cells[idx]
                        if isinstance(v, str) and v.strip().startswith('='):
                            continue
                        num = float(v)
                        if num != 0:
                            ef_value = num
                            ef_idx = idx
                            break
                    except (ValueError, TypeError):
                        continue

                unit = ''
                emission_source_reference = ''
                for idx in range(5, 11):
                    v = cells[idx]
                    if isinstance(v, (int, float)):
                        continue
                    if not unit and looks_like_unit(v):
                        unit = v
                    if not emission_source_reference and looks_like_reference(v):
                        emission_source_reference = v

                if ef_idx is not None:
                    if not unit:
                        candidate = cells.get(ef_idx + 1)
                        candidate_str = str(candidate).strip() if candidate is not None else ''
                        if candidate_str and not looks_like_reference(candidate) and not isinstance(candidate, (int, float)):
                            unit = candidate
                    if not emission_source_reference:
                        candidate = cells.get(ef_idx + 2)
                        candidate_str = str(candidate).strip() if candidate is not None else ''
                        if candidate_str and not looks_like_unit(candidate) and not isinstance(candidate, (int, float)):
                            emission_source_reference = candidate

                if not unit:
                    for idx in range(5, 11):
                        v = cells[idx]
                        if isinstance(v, str) and 'co2e' in v.lower():
                            unit = v
                            break

                if not emission_source_reference:
                    for idx in range(10, 4, -1):
                        v = cells[idx]
                        if isinstance(v, str) and v.strip() and looks_like_reference(v):
                            emission_source_reference = v
                            break
                if not emission_source_reference:
                    for idx in range(5, 11):
                        v = cells[idx]
                        if isinstance(v, str) and 'ecoinvent' in v.lower():
                            emission_source_reference = extract_ecoinvent_reference(v)
                            if emission_source_reference:
                                break

                geography = ''
            else:
                geography = self.safe_get_cell(row, 5)
                cells = {i: self.safe_get_cell(row, i) for i in range(6, 12)}

                ef_value = 0.0
                ef_idx = None
                for idx in [6, 7, 8]:
                    try:
                        v = cells[idx]
                        if isinstance(v, str) and v.strip().startswith('='):
                            continue
                        num = float(v)
                        if num != 0:
                            ef_value = num
                            ef_idx = idx
                            break
                    except (ValueError, TypeError):
                        continue

                unit = ''
                emission_source_reference = ''
                for idx in range(6, 12):
                    v = cells[idx]
                    if isinstance(v, (int, float)):
                        continue
                    if not unit and looks_like_unit(v):
                        unit = v
                    if not emission_source_reference and looks_like_reference(v):
                        emission_source_reference = v

                if ef_idx is not None:
                    if not unit:
                        candidate = cells.get(ef_idx + 1)
                        candidate_str = str(candidate).strip() if candidate is not None else ''
                        if candidate_str and not looks_like_reference(candidate) and not isinstance(candidate, (int, float)):
                            unit = candidate
                    if not emission_source_reference:
                        candidate = cells.get(ef_idx + 2)
                        candidate_str = str(candidate).strip() if candidate is not None else ''
                        if candidate_str and not looks_like_unit(candidate) and not isinstance(candidate, (int, float)):
                            emission_source_reference = candidate

                if not unit:
                    for idx in range(6, 12):
                        v = cells[idx]
                        if isinstance(v, str) and 'co2e' in v.lower():
                            unit = v
                            break

                if not emission_source_reference:
                    for idx in range(11, 5, -1):
                        v = cells[idx]
                        if isinstance(v, str) and v.strip() and looks_like_reference(v):
                            emission_source_reference = v
                            break
                if not emission_source_reference:
                    for idx in range(6, 12):
                        v = cells[idx]
                        if isinstance(v, str) and 'ecoinvent' in v.lower():
                            emission_source_reference = extract_ecoinvent_reference(v)
                            if emission_source_reference:
                                break

            return {
                'number': number,
                'category': category,
                'emission_source': emission_source,
                'facility': activity_name,
                'activity_name': activity_name,
                'geography': geography,
                'ncv': 0,
                'unit': unit,
                'ox_rate': 0,
                'CO2_emission_cv_factor': 0,
                'CH4_emission_cv_factor': 0,
                'N2O_emission_cv_factor': 0,
                'CO2_emission_factor': ef_value,
                'CH4_emission_factor': 0,
                'N2O_emission_factor': 0,
                'emission_source_reference': emission_source_reference,
            }
        except Exception:
            return {}


__all__ = ['EmissionFactorReader']
