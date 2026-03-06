# -*- coding: utf-8 -*-
"""
重构 data_reader.py 的脚本
"""
import re


def refactor_data_reader():
    # Read the current data_reader.py
    with open('data_reader.py', 'r', encoding='utf-8') as f:
        content = f.read()

    print("Starting refactoring...")

    # 1. Add import for ReportConfig at the top
    if 'from report_config import ReportConfig' not in content:
        content = content.replace(
            'import openpyxl \nimport csv\nimport os',
            'import openpyxl \nimport csv\nimport os\nfrom report_config import ReportConfig\n'
        )
        print('[1/7] Added import for ReportConfig')
    else:
        print('[1/7] Import already exists')

    # 2. Update __init__ to include company_name and reporting_period
    if 'self.company_name = None' not in content:
        content = content.replace(
            'self.file_type = None',
            'self.file_type = None\n        self.company_name = None\n        self.reporting_period = \'2024年\'  # 默认报告期'
        )
        print('[2/7] Updated __init__ to include company_name and reporting_period')
    else:
        print('[2/7] __init__ already updated')

    # 3. Remove the get_quantification_methods method
    pattern1 = r"    def get_quantification_methods\(self\):\n        \"\"\"\n        返回各类别的量化方法说明\n        用于动态生成\"量化方法说明\"章节\n        \"\"\"\n        return \{[^}]+\}"
    before = content
    content = re.sub(pattern1, '', content, flags=re.DOTALL)
    if content != before:
        print('[3/7] Removed get_quantification_methods method')
    else:
        print('[3/7] get_quantification_methods already removed or not found')

    # 4. Remove the get_scope_3_category_name method
    pattern2 = r"    def get_scope_3_category_name\(self, category_num\):\n        \"\"\"获取范围三全部 15 个类别名称\"\"\"\n        names = \{[^}]+\}\n        return names\.get\(category_num, f\"类别\{category_num\}\"\)"
    before = content
    content = re.sub(pattern2, '', content, flags=re.DOTALL)
    if content != before:
        print('[4/7] Removed get_scope_3_category_name method')
    else:
        print('[4/7] get_scope_3_category_name already removed or not found')

    # 5. Replace quantification_methods calls
    old_call = "result['quantification_methods'] = self.get_quantification_methods()"
    new_call = """# Use ReportConfig for quantification methods
            report_config = ReportConfig(
                company_name or '某公司',
                reporting_period or '2024年'
            )
            result['quantification_methods'] = report_config.get_quantification_methods()"""
    content = content.replace(old_call, new_call)
    print('[5/7] Updated quantification_methods calls')

    # 6. Replace scope_3_category_names calls
    old_names = """result['scope_3_category_names'] = {}
            for i in range(1, 16):
                result['scope_3_category_names'][f'category_{i}'] = self.get_scope_3_category_name(i)"""
    new_names = """# Use ReportConfig for scope 3 category names
            report_config_names = ReportConfig()
            result['scope_3_category_names'] = report_config_names.get_all_scope_3_category_names()"""
    content = content.replace(old_names, new_names)

    old_names2 = """data['scope_3_category_names'] = {}
        for i in range(1, 16):
            data['scope_3_category_names'][f'category_{i}'] = self.get_scope_3_category_name(i)"""
    new_names2 = """# Use ReportConfig for scope 3 category names
        report_config_names = ReportConfig()
        data['scope_3_category_names'] = report_config_names.get_all_scope_3_category_names()"""
    content = content.replace(old_names2, new_names2)
    print('[6/7] Updated scope_3_category_names calls')

    # 7. Add self.company_name and self.reporting_period assignment
    find_text = "report_year = '2024'  # 直接提取年份"
    replace_text = """report_year = '2024'  # 直接提取年份

        # 赋值给 self 变量，供 ReportConfig 使用
        self.company_name = company_name
        self.reporting_period = report_period or '2024年'  # 默认值"""

    if find_text in content and 'self.company_name = company_name' not in content:
        content = content.replace(find_text, replace_text)
        print('[7/7] Added self.company_name and self.reporting_period assignment')
    else:
        print('[7/7] Already has self.company_name assignment or find_text not found')

    # Also replace the other occurrence
    find_text2 = "data['report_year'] = year_match.group(1) if year_match else '2024'"
    replace_text2 = """data['report_year'] = year_match.group(1) if year_match else '2024'

            # 赋值给 self 变量，供 ReportConfig 使用
            self.company_name = data.get('company_name')
            self.reporting_period = data.get('reporting_period') or '2024年'"""

    if find_text2 in content:
        # Only replace if we haven't already added this
        if content.count('self.company_name = data.get') < 1:
            content = content.replace(find_text2, replace_text2)
            print('[7/7] Added self.company_name assignment in extract_data')

    # Write the updated content
    with open('data_reader.py', 'w', encoding='utf-8') as f:
        f.write(content)

    print('\n=== Refactoring complete! ===')


if __name__ == '__main__':
    refactor_data_reader()
