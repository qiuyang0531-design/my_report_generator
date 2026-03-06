# -*- coding: utf-8 -*-
"""
修复format_emission函数，确保0值返回空字符串
"""
import re

def fix_data_reader():
    """修复data_reader.py中的format_emission函数"""
    with open('data_reader.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # 查找并替换format_emission函数
    # 旧的模式：返回'0.00'
    old_pattern = r"def format_emission\(val\):.*?return '0\.00'"
    new_function = '''def format_emission(val):
                        if val is None or val == 0:
                            return ''
                        try:
                            float_value = float(val)
                            if float_value == 0:
                                return ''
                            return f"{float_value:.2f}"
                        except (ValueError, TypeError):
                            return '''''

    # 使用正则表达式替换（DOTALL以匹配多行）
    new_content = re.sub(old_pattern, new_function, content, flags=re.DOTALL)

    if new_content != content:
        with open('data_reader.py', 'w', encoding='utf-8') as f:
            f.write(new_content)
        print('[OK] 已修复data_reader.py中的format_emission函数')
        return True
    else:
        print('[SKIP] data_reader.py中的format_emission已经是正确的')
        return False

if __name__ == '__main__':
    fix_data_reader()
