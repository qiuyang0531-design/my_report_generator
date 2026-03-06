"""
重构 data_reader.py 的脚本
1. 添加 ReportConfig 导入
2. 移除旧的 get_quantification_methods 和 get_scope_3_category_name 方法
3. 更新调用这些方法的地方
4. 确保 company_name 和 reporting_period 被赋值给 self
"""
import re


def refactor_data_reader():
    # Read the current data_reader.py
    with open('data_reader.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Add import for ReportConfig at the top
    if 'from report_config import ReportConfig' not in content:
        content = content.replace(
            'import openpyxl \nimport csv\nimport os',
            'import openpyxl \nimport csv\nimport os\nfrom report_config import ReportConfig\n'
        )
        print('✅ Added import for ReportConfig')

    # 2. Update __init__ to include company_name and reporting_period
    old_init = '''    def __init__(self, filepath):
        """
        初始化时，加载 Excel 工作簿。
        """
        self.workbook = None
        self.filepath = filepath
        self.file_type = None'''

    new_init = '''    def __init__(self, filepath):
        """
        初始化时，加载 Excel 工作簿。
        """
        self.workbook = None
        self.filepath = filepath
        self.file_type = None
        self.company_name = None
        self.reporting_period = '2024年'  # 默认报告期'''

    if old_init in content:
        content = content.replace(old_init, new_init)
        print('✅ Updated __init__ to include company_name and reporting_period')

    # 3. Remove the get_quantification_methods method
    quant_pattern = r'''    def get_quantification_methods\(self\):
    """
    返回各类别的量化方法说明
    用于动态生成"量化方法说明"章节
    """
    return \{[^}]+\}'''

    content_before = content
    content = re.sub(quant_pattern, '', content, flags=re.DOTALL)
    if content != content_before:
        print('✅ Removed get_quantification_methods method')

    # 4. Remove the get_scope_3_category_name method
    name_pattern = r'''    def get_scope_3_category_name\(self, category_num\):
        """获取范围三全部 15 个类别名称"""
        names = \{[^}]+\}
        return names\.get\(category_num, f"类别\{category_num\}"\)'''

    content_before = content
    content = re.sub(name_pattern, '', content, flags=re.DOTALL)
    if content != content_before:
        print('✅ Removed get_scope_3_category_name method')

    # 5. Replace quantification_methods calls in extract_data (around line 761)
    old_quant_call_1 = '''            result['quantification_methods'] = self.get_quantification_methods()'''
    new_quant_call_1 = '''            # Use ReportConfig for quantification methods
            report_config = ReportConfig(
                company_name or '某公司',
                reporting_period or '2024年'
            )
            result['quantification_methods'] = report_config.get_quantification_methods()'''

    if old_quant_call_1 in content:
        content = content.replace(old_quant_call_1, new_quant_call_1)
        print('✅ Updated quantification_methods calls in extract_data')

    # 6. Replace quantification_methods calls in extract_data_from_xlsx_dynamic (around line 1656)
    old_quant_call_2 = '''        data['quantification_methods'] = self.get_quantification_methods()'''
    new_quant_call_2 = '''        # Use ReportConfig for quantification methods
        report_config = ReportConfig(
            self.company_name or '某公司',
            self.reporting_period or '2024年'
        )
        data['quantification_methods'] = report_config.get_quantification_methods()'''

    # Replace the second occurrence (in extract_data_from_xlsx_dynamic)
    if content.count(old_quant_call_2) >= 1:
        # Replace only the last occurrence (in extract_data_from_xlsx_dynamic)
        parts = content.rsplit(old_quant_call_2, 1)
        if len(parts) > 1:
            content = parts[0] + new_quant_call_2 + parts[1]
            print('✅ Updated quantification_methods calls in extract_data_from_xlsx_dynamic')

    # 7. Replace scope_3_category_names calls
    old_names_call_1 = '''            result['scope_3_category_names'] = {}
            for i in range(1, 16):
                result['scope_3_category_names'][f'category_{i}'] = self.get_scope_3_category_name(i)'''

    new_names_call_1 = '''            # Use ReportConfig for scope 3 category names
            report_config_names = ReportConfig()
            result['scope_3_category_names'] = report_config_names.get_all_scope_3_category_names()'''

    if old_names_call_1 in content:
        content = content.replace(old_names_call_1, new_names_call_1)
        print('✅ Updated scope_3_category_names calls in extract_data')

    old_names_call_2 = '''        data['scope_3_category_names'] = {}
        for i in range(1, 16):
            data['scope_3_category_names'][f'category_{i}'] = self.get_scope_3_category_name(i)'''

    new_names_call_2 = '''        # Use ReportConfig for scope 3 category names
        report_config_names = ReportConfig()
        data['scope_3_category_names'] = report_config_names.get_all_scope_3_category_names()'''

    if old_names_call_2 in content:
        content = content.replace(old_names_call_2, new_names_call_2)
        print('✅ Updated scope_3_category_names calls in extract_data_from_xlsx_dynamic')

    # 8. Add self.company_name and self.reporting_period assignment in extract_data_from_xlsx_dynamic
    # Find the section where company_name is set and add self.company_name = company_name
    old_company_assign = '''        company_name = self.find_value_by_label(main_sheet, '组织名称：')
        report_period = self.find_value_by_label(main_sheet, '盘查覆盖周期:')
        # 从报告周期中提取年份（假设格式为"2024年1月1日至2024年12月31日"）
        report_year = '2024'  # 直接提取年份'''

    new_company_assign = '''        company_name = self.find_value_by_label(main_sheet, '组织名称：')
        report_period = self.find_value_by_label(main_sheet, '盘查覆盖周期:')
        # 从报告周期中提取年份（假设格式为"2024年1月1日至2024年12月31日"）
        report_year = '2024'  # 直接提取年份

        # 赋值给 self 变量，供 ReportConfig 使用
        self.company_name = company_name
        self.reporting_period = report_period or '2024年'  # 默认值'

    if old_company_assign in content:
        content = content.replace(old_company_assign, new_company_assign)
        print('✅ Added self.company_name and self.reporting_period assignment')

    # Write the updated content
    with open('data_reader.py', 'w', encoding='utf-8') as f:
        f.write(content)

    print('\n🎉 data_reader.py refactored successfully!')


if __name__ == '__main__':
    refactor_data_reader()
