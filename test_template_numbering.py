# -*- coding: utf-8 -*-
"""
Test script to verify Jinja2 numbering works correctly
"""
from jinja2 import Environment

# Test the template logic directly
env = Environment()

def to_chinese_num(n):
    """Convert number to Chinese numeral"""
    chinese_map = {
        1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
        6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
        11: '十一', 12: '十二'
    }
    return chinese_map.get(n, str(n))

env.filters['cn_num'] = to_chinese_num

# Template with namespace counter
template_str = """
{% set ns = namespace(count=1) %}
{% for item in items -%}
（{{ ns.count }}）Item: {{ item.name }}
{% set ns.count = ns.count + 1 %}
{%- endfor %}
"""

template = env.from_string(template_str)

# Test data
items = [
    {'name': 'First'},
    {'name': 'Second'},
    {'name': 'Third'},
]

result = template.render(items=items)
print("=== Template Rendering Result ===")
print(result)
print("=== End ===")

# Also test with cn_nums dictionary
template_str2 = """
{% set ns = namespace(count=1) %}
{% set cn_nums = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五'} %}
{% for item in items -%}
（{{ cn_nums[ns.count] }}）Item: {{ item.name }}
{% set ns.count = ns.count + 1 %}
{%- endfor %}
"""

template2 = env.from_string(template_str2)
result2 = template2.render(items=items)
print("\n=== With cn_nums dictionary ===")
print(result2)
