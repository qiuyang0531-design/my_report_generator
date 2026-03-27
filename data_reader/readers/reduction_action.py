"""
减排措施读取模块
=====================

从Excel中提取今年已实施和明年计划的减排措施数据。
"""

import re
from typing import Dict, Any, List
from .base import BaseReader


class ReductionActionReader(BaseReader):
    """减排措施读取器"""

    def extract(self) -> Dict[str, Any]:
        """
        提取减排措施数据，区分今年已实施和明年计划

        Returns:
            包含减排措施数据的字典
        """
        result = {
            'implemented_reduction_items': [],  # 今年已实施
            'planned_reduction_items': [],       # 明年计划
        }

        # 查找减排措施统计表（遍历所有Sheet）
        reduction_sheet = None
        for sheet in self.workbook.worksheets:
            sheet_name = sheet.title
            # 检查是否包含相关关键词
            if any(keyword in sheet_name for keyword in ['减排', '措施', '节能', '统计']):
                reduction_sheet = sheet
                print(f"[减排措施] 找到表: {sheet_name}")
                break

        if not reduction_sheet:
            print("[减排措施] 未找到减排措施表")
            return result

        self._parse_reduction_sheet(reduction_sheet, result)

        return result

    def _parse_reduction_sheet(self, ws, result: Dict[str, Any]):
        """解析减排措施表"""
        current_section = None

        # 初始化默认的列名映射（假设标准格式）
        default_mapping = {
            '序号': 0,
            'number': 0,
            '措施名称': 2,
            'plan_name': 2,
            '方案名称': 2,
            '实施方案名称': 2,
            '措施描述': 3,
            'description': 3,
            '方案描述': 3,
            '负责部门': 4,
            'department': 4,
            '实施单位': 4,
            '实施进度': 5,
            'progress': 5,
            '完成情况': 5,
            '预计减排量': 6,
            'estimated_savings': 6,
            '预计节能量': 6,
        }

        # 遍历所有行
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            # 跳过完全空的行
            if not any(row):
                continue

            # 获取第一个非空单元格的值
            first_cell_value = None
            for cell in row:
                if cell:
                    cell_str = str(cell).strip()
                    # 跳过Jinja2模板标签
                    if cell_str.startswith(('{%', '{{', '{#')):
                        continue
                    if not first_cell_value:
                        first_cell_value = cell_str

            if not first_cell_value:
                continue

            # 检查所有列是否包含年份表头（如 "2024年已实施的减排措施"）
            year_match = None
            for cell in row:
                if cell:
                    cell_str = str(cell).strip()
                    # 匹配：已实施、拟实施、计划实施
                    year_match = re.search(r'(\d{4})年.*?(已实施|拟实施|计划实施)', cell_str)
                    if year_match:
                        break
            if year_match:
                current_year = year_match.group(1)
                status = year_match.group(2)

                if '已实施' in status:
                    current_section = 'implemented'
                    print(f"[减排措施] 发现{current_year}年已实施部分 (行{row_idx})")
                elif '拟实施' in status or '计划实施' in status:
                    current_section = 'planned'
                    print(f"[减排措施] 发现{current_year}年计划实施部分 (行{row_idx})")
                continue

            # 检查是否是表头行
            if first_cell_value in ['序号', '编号', '措施名称', '方案名称', '项目', '实施方案']:
                continue

            # 如果没有确定section，输出调试信息
            if not current_section:
                print(f"[减排措施] 跳过行{row_idx} (未确定section): {first_cell_value[:30]}")
                continue

            # 解析减排措施条目
            item = self._parse_reduction_item(row, default_mapping)
            if item:
                if current_section == 'implemented':
                    result['implemented_reduction_items'].append(item)
                    print(f"[减排措施] 添加已实施条目: {item.get('plan_name', '')[:30]}")
                elif current_section == 'planned':
                    result['planned_reduction_items'].append(item)
                    print(f"[减排措施] 添加计划条目: {item.get('plan_name', '')[:30]}")
            else:
                print(f"[减排措施] 行{row_idx}解析失败: {first_cell_value[:30]}")

        print(f"[减排措施] 今年已实施: {len(result['implemented_reduction_items'])} 条")
        print(f"[减排措施] 明年计划: {len(result['planned_reduction_items'])} 条")

    def _extract_header_mapping(self, header_row, mapping: Dict[str, int]):
        """提取表头列名映射"""
        header_mapping = {
            '序号': 0,
            '编号': 0,
            'number': 0,
            '项目': 0,
            '措施名称': 0,
            'plan_name': 0,
            '方案名称': 0,
            '实施方案名称': 0,
            '措施描述': 0,
            'description': 0,
            '方案描述': 0,
            '负责部门': 0,
            'department': 0,
            '实施单位': 0,
            '实施进度': 0,
            'progress': 0,
            '完成情况': 0,
            '预计减排量': 0,
            'estimated_savings': 0,
            '预计节能量': 0,
        }

        for col_idx, cell_value in enumerate(header_row):
            if cell_value:
                cell_str = str(cell_value).strip()
                for key in header_mapping:
                    if key in cell_str:
                        mapping[key] = col_idx
                        break

    def _parse_reduction_item(self, row, mapping: Dict[str, int]) -> Dict[str, Any]:
        """解析单条减排措施"""
        def get_value(*keys):
            """按优先级获取值"""
            for key in keys:
                col_idx = mapping.get(key, -1)
                if col_idx >= 0 and col_idx < len(row):
                    val = row[col_idx]
                    if val:
                        val_str = str(val).strip()
                        # 跳过Jinja2模板标签
                        if not val_str.startswith(('{%', '{{', '{#')):
                            return val_str
            return ""

        item = {
            'plan_name': get_value('措施名称', '方案名称', '实施方案名称', '项目'),
            'description': get_value('措施描述', '方案描述', '详情'),
            'department': get_value('负责部门', '实施单位', '部门'),
            'progress': get_value('实施进度', '完成情况', '进度'),
            'estimated_savings': get_value('预计减排量', '预计节能量', '减排效果'),
        }

        # 过滤掉全空的条目
        if not any(item.values()):
            return None

        # 确保plan_name不为空
        if not item['plan_name']:
            return None

        return item


__all__ = ['ReductionActionReader']
