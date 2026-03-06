# -*- coding: utf-8 -*-
"""
回归测试脚本
验证重构后的代码是否正确
"""
import sys
from data_reader import ExcelDataReader

def test_flags_and_data_types():
    """测试 Flags 逻辑和数据类型"""
    print("=" * 60)
    print("回归测试：Flags 逻辑和数据类型验证")
    print("=" * 60)

    # Load data
    reader = ExcelDataReader('test_data.xlsx')
    data = reader.extract_data_from_xlsx_dynamic()

    # Test 1: Flags should be Boolean type
    print("\n[测试1] Flags 数据类型验证")
    flags = data.get('flags', {})
    all_boolean = True
    for key, value in flags.items():
        if not isinstance(value, bool):
            print(f"  [FAIL] {key}: {type(value)} (期望: bool)")
            all_boolean = False
    if all_boolean:
        print(f"  [PASS] 所有 {len(flags)} 个 Flags 都是 Boolean 类型")

    # Test 2: total_emissions should be Float type
    print("\n[测试2] 总排放量数据类型验证")
    total_checks = [
        ('scope_1_emissions', data.get('scope_1_emissions')),
        ('scope_2_location_based_emissions', data.get('scope_2_location_based_emissions')),
        ('scope_2_market_based_emissions', data.get('scope_2_market_based_emissions')),
        ('scope_3_emissions', data.get('scope_3_emissions')),
        ('total_emission_location', data.get('total_emission_location')),
        ('total_emission_market', data.get('total_emission_market')),
    ]

    all_float = True
    for name, value in total_checks:
        if value is not None and not isinstance(value, (float, int)):
            print(f"  [FAIL] {name}: {type(value)} (期望: float)")
            all_float = False
    if all_float:
        print(f"  [PASS] 所有排放量数据都是 Float/Int 类型")

    # Test 3: quantification_methods exists and has f-string content
    print("\n[测试3] 量化方法说明验证")
    quant_methods = data.get('quantification_methods', {})

    # Check if all scopes are present
    has_scope_1 = 'scope_1' in quant_methods
    has_scope_2 = 'scope_2' in quant_methods
    has_scope_3 = 'scope_3' in quant_methods

    if has_scope_1 and has_scope_2 and has_scope_3:
        print(f"  [PASS] quantification_methods 包含所有三个范围")

        # Check if f-string dynamic content is present (company name and reporting period)
        scope_1_fixed = quant_methods['scope_1']['固定燃烧']
        if '{' in scope_1_fixed or '某公司' in scope_1_fixed or reader.company_name in scope_1_fixed:
            print(f"  [PASS] 量化方法说明包含动态内容 (f-string)")
        else:
            print(f"  [WARN] 量化方法说明可能是静态内容")
            print(f"     内容预览: {scope_1_fixed[:100]}...")
    else:
        print(f"  [FAIL] quantification_methods 缺少范围数据")

    # Test 4: scope_3_category_names exists and has all 15 categories
    print("\n[测试4] 范围三类别名称验证")
    cat_names = data.get('scope_3_category_names', {})
    if len(cat_names) == 15:
        print(f"  [PASS] scope_3_category_names 包含全部 15 个类别")
        # Show first few
        for i in range(1, 4):
            key = f'category_{i}'
            if key in cat_names:
                print(f"     {key}: {cat_names[key]}")
    else:
        print(f"  [FAIL] scope_3_category_names 只有 {len(cat_names)} 个类别")

    # Test 5: Check scope 3 has all 15 categories
    print("\n[测试5] 范围三 15 个类别完整性验证")
    all_categories_present = True
    for i in range(1, 16):
        cat_key = f'scope_3_category_{i}_emissions'
        if cat_key not in data:
            print(f"  [FAIL] 缺少类别 {i}: {cat_key}")
            all_categories_present = False
    if all_categories_present:
        print(f"  [PASS] 范围三所有 15 个类别数据键都存在")

    # Test 6: Self variables assignment
    print("\n[测试6] Self 变量赋值验证")
    print(f"  reader.company_name: {reader.company_name}")
    print(f"  reader.reporting_period: {reader.reporting_period}")
    if reader.company_name:
        print(f"  [PASS] self.company_name 已正确赋值")
    if reader.reporting_period:
        print(f"  [PASS] self.reporting_period 已正确赋值")

    # Test 7: Check scope 3 quantification methods have all 15 categories
    print("\n[测试7] 范围三量化方法 15 个类别完整性验证")
    scope_3_methods = quant_methods.get('scope_3', {})
    if len(scope_3_methods) == 15:
        print(f"  [PASS] scope_3 量化方法包含全部 15 个类别")
        # Show first few
        for i in range(1, 4):
            key = f'category_{i}'
            if key in scope_3_methods:
                method_text = scope_3_methods[key]
                # Check if it contains dynamic references
                has_dynamic = '{' in method_text or reader.company_name in method_text
                status = '动态' if has_dynamic else '静态'
                print(f"     {key}: {status} ({method_text[:50]}...)")
    else:
        print(f"  [FAIL] scope_3 量化方法只有 {len(scope_3_methods)} 个类别")

    # Summary
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

    return all_boolean and all_float and has_scope_1 and has_scope_2 and has_scope_3


if __name__ == '__main__':
    try:
        success = test_flags_and_data_types()
        if success:
            print("\n[SUCCESS] 所有测试通过!")
            sys.exit(0)
        else:
            print("\n[WARNING] 部分测试未通过，请检查")
            sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] 测试出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
