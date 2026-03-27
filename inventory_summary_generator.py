"""
基准年温室气体清单汇总数据生成器
=====================================

用于生成两张汇总表的完整上下文数据：
1. （一）基准年温室气体清单（基于位置）
2. （二）基准年温室气体清单（基于市场）
"""

from typing import Dict, Any, Optional


def format_number(value) -> str:
    """
    格式化数值：保留两位小数，千分位逗号

    Args:
        value: 输入值（数字、字符串等）

    Returns:
        格式化后的字符串，如 '1,069,378.75' 或 '0.00'
    """
    if value is None or value == '':
        return '0.00'
    try:
        f_value = float(value)
        return f"{f_value:,.2f}"
    except (ValueError, TypeError):
        return '0.00'


def generate_inventory_context(raw_data: Dict[str, Any], green_power_data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    生成基准年温室气体清单的汇总数据

    数据结构：
    - summary_data['scope1']: 范围一数据（共享）
    - summary_data['scope3']: 范围三数据（共享），包含15个类别
    - summary_data['scope2']['loc']: 范围二基于位置
    - summary_data['scope2']['mar']: 范围二基于市场
    - summary_data['total_loc']: 企业总计（基于位置）
    - summary_data['total_mar']: 企业总计（基于市场）

    Args:
        raw_data: 原始排放数据，包含各范围排放源
                  期望字段：scope1_co2, scope1_ch4, ..., scope1_total
                            cat1_co2, cat1_ch4, ..., cat1_total (类别1-15)
                            scope2_co2, scope2_ch4, ..., scope2_total
                            scope3_co2, scope3_ch4, ..., scope3_total
        green_power_data: 绿电/绿证数据（可选，用于基于市场计算）
                          期望字段：gec_mwh (绿证GEC电量), green_elec_mwh (绿电), other_mwh (其他)

    Returns:
        summary_data: 符合指定结构的字典，所有数值已格式化
    """
    summary_data = {}

    # ============================================================
    # 1. 共享部分：范围一排放数据
    # ============================================================
    summary_data['scope1'] = {
        'co2': format_number(raw_data.get('scope1_co2', 0)),
        'ch4': format_number(raw_data.get('scope1_ch4', 0)),
        'n2o': format_number(raw_data.get('scope1_n2o', 0)),
        'hfcs': format_number(raw_data.get('scope1_hfcs', 0)),
        'pfcs': format_number(raw_data.get('scope1_pfcs', 0)),
        'sf6': format_number(raw_data.get('scope1_sf6', 0)),
        'nf3': format_number(raw_data.get('scope1_nf3', 0)),
        'total': format_number(raw_data.get('scope1_total', 0)),
    }

    # ============================================================
    # 2. 共享部分：范围三排放数据（包含15个类别明细）
    # ============================================================
    scope3_categories = {}
    scope3_totals = {
        'co2': 0.0, 'ch4': 0.0, 'n2o': 0.0,
        'hfcs': 0.0, 'pfcs': 0.0, 'sf6': 0.0, 'nf3': 0.0, 'total': 0.0
    }

    for i in range(1, 16):
        cat_key = f'cat{i}'
        category = raw_data.get(cat_key, {})

        cat_co2 = float(category.get('co2', 0))
        cat_ch4 = float(category.get('ch4', 0))
        cat_n2o = float(category.get('n2o', 0))
        cat_hfcs = float(category.get('hfcs', 0))
        cat_pfcs = float(category.get('pfcs', 0))
        cat_sf6 = float(category.get('sf6', 0))
        cat_nf3 = float(category.get('nf3', 0))
        cat_total = float(category.get('total', 0))

        # 累加到范围三总计
        scope3_totals['co2'] += cat_co2
        scope3_totals['ch4'] += cat_ch4
        scope3_totals['n2o'] += cat_n2o
        scope3_totals['hfcs'] += cat_hfcs
        scope3_totals['pfcs'] += cat_pfcs
        scope3_totals['sf6'] += cat_sf6
        scope3_totals['nf3'] += cat_nf3
        scope3_totals['total'] += cat_total

        # 存储类别数据
        scope3_categories[cat_key] = {
            'co2': format_number(cat_co2),
            'ch4': format_number(cat_ch4),
            'n2o': format_number(cat_n2o),
            'hfcs': format_number(cat_hfcs),
            'pfcs': format_number(cat_pfcs),
            'sf6': format_number(cat_sf6),
            'nf3': format_number(cat_nf3),
            'total': format_number(cat_total),
        }

    # 添加格式化的范围三总计 - 平铺到scope3字典中（与cat1, cat2等平级）
    # 模板使用 {{summary_data.scope3.co2}} 而不是 {{summary_data.scope3.total.co2}}
    scope3_categories['co2'] = format_number(scope3_totals['co2'])
    scope3_categories['ch4'] = format_number(scope3_totals['ch4'])
    scope3_categories['n2o'] = format_number(scope3_totals['n2o'])
    scope3_categories['hfcs'] = format_number(scope3_totals['hfcs'])
    scope3_categories['pfcs'] = format_number(scope3_totals['pfcs'])
    scope3_categories['sf6'] = format_number(scope3_totals['sf6'])
    scope3_categories['nf3'] = format_number(scope3_totals['nf3'])
    scope3_categories['total'] = format_number(scope3_totals['total'])
    summary_data['scope3'] = scope3_categories

    # ============================================================
    # 3. 处理绿电数据（用于基于市场计算）
    # ============================================================
    if green_power_data is None:
        green_power_data = {}

    gec_mwh = float(green_power_data.get('gec_mwh', 0))      # 绿证GEC电量 (MWh)
    green_elec_mwh = float(green_power_data.get('green_elec_mwh', 0))  # 绿电 (MWh)
    other_mwh = float(green_power_data.get('other_mwh', 0))  # 其他 (MWh)

    # 绿电数据汇总
    summary_data['green_power'] = {
        'gec_mwh': format_number(gec_mwh),
        'gec_kwh': format_number(gec_mwh * 1000),  # MWh -> kWh
        'green_elec_mwh': format_number(green_elec_mwh),
        'green_elec_kwh': format_number(green_elec_mwh * 1000),
        'other_mwh': format_number(other_mwh),
        'other_kwh': format_number(other_mwh * 1000),
        'total_mwh': format_number(gec_mwh + green_elec_mwh + other_mwh),
        'total_kwh': format_number((gec_mwh + green_elec_mwh + other_mwh) * 1000),
    }

    # ============================================================
    # 4. 范围二数据（基于位置 vs 基于市场）
    # ============================================================
    # 直接从 raw_data 中读取两组范围二数据
    # 基于位置的数据
    scope2_loc_co2 = float(raw_data.get('scope2_loc_co2', 0))
    scope2_loc_ch4 = float(raw_data.get('scope2_loc_ch4', 0))
    scope2_loc_n2o = float(raw_data.get('scope2_loc_n2o', 0))
    scope2_loc_hfcs = float(raw_data.get('scope2_loc_hfcs', 0))
    scope2_loc_pfcs = float(raw_data.get('scope2_loc_pfcs', 0))
    scope2_loc_sf6 = float(raw_data.get('scope2_loc_sf6', 0))
    scope2_loc_nf3 = float(raw_data.get('scope2_loc_nf3', 0))
    scope2_loc_total = float(raw_data.get('scope2_loc_total', 0))

    # 基于市场的数据
    scope2_mar_co2 = float(raw_data.get('scope2_mar_co2', 0))
    scope2_mar_ch4 = float(raw_data.get('scope2_mar_ch4', 0))
    scope2_mar_n2o = float(raw_data.get('scope2_mar_n2o', 0))
    scope2_mar_hfcs = float(raw_data.get('scope2_mar_hfcs', 0))
    scope2_mar_pfcs = float(raw_data.get('scope2_mar_pfcs', 0))
    scope2_mar_sf6 = float(raw_data.get('scope2_mar_sf6', 0))
    scope2_mar_nf3 = float(raw_data.get('scope2_mar_nf3', 0))
    scope2_mar_total = float(raw_data.get('scope2_mar_total', 0))

    # 基于位置：使用原始数据
    summary_data['scope2'] = {}
    summary_data['scope2']['loc'] = {
        'co2': format_number(scope2_loc_co2),
        'ch4': format_number(scope2_loc_ch4),
        'n2o': format_number(scope2_loc_n2o),
        'hfcs': format_number(scope2_loc_hfcs),
        'pfcs': format_number(scope2_loc_pfcs),
        'sf6': format_number(scope2_loc_sf6),
        'nf3': format_number(scope2_loc_nf3),
        'total': format_number(scope2_loc_total),
    }

    # 基于市场：使用传入的市场数据
    summary_data['scope2']['mkt'] = {
        'co2': format_number(scope2_mar_co2),
        'ch4': format_number(scope2_mar_ch4),
        'n2o': format_number(scope2_mar_n2o),
        'hfcs': format_number(scope2_mar_hfcs),
        'pfcs': format_number(scope2_mar_pfcs),
        'sf6': format_number(scope2_mar_sf6),
        'nf3': format_number(scope2_mar_nf3),
        'total': format_number(scope2_mar_total),
    }

    # 保持 mar 键作为 mkt 的别名（向后兼容）
    summary_data['scope2']['mar'] = summary_data['scope2']['mkt']

    # ============================================================
    # 5. 企业总计（基于位置 vs 基于市场）
    # ============================================================
    scope3_co2 = scope3_totals['co2']
    scope3_ch4 = scope3_totals['ch4']
    scope3_n2o = scope3_totals['n2o']
    scope3_hfcs = scope3_totals['hfcs']
    scope3_pfcs = scope3_totals['pfcs']
    scope3_sf6 = scope3_totals['sf6']
    scope3_nf3 = scope3_totals['nf3']
    scope3_total = scope3_totals['total']

    scope1_co2 = float(raw_data.get('scope1_co2', 0))
    scope1_ch4 = float(raw_data.get('scope1_ch4', 0))
    scope1_n2o = float(raw_data.get('scope1_n2o', 0))
    scope1_hfcs = float(raw_data.get('scope1_hfcs', 0))
    scope1_pfcs = float(raw_data.get('scope1_pfcs', 0))
    scope1_sf6 = float(raw_data.get('scope1_sf6', 0))
    scope1_nf3 = float(raw_data.get('scope1_nf3', 0))
    scope1_total = float(raw_data.get('scope1_total', 0))

    # 基于位置总计 = 范围一 + 范围二(位置) + 范围三
    summary_data['total_loc'] = {
        'co2': format_number(scope1_co2 + scope2_loc_co2 + scope3_co2),
        'ch4': format_number(scope1_ch4 + scope2_loc_ch4 + scope3_ch4),
        'n2o': format_number(scope1_n2o + scope2_loc_n2o + scope3_n2o),
        'hfcs': format_number(scope1_hfcs + scope2_loc_hfcs + scope3_hfcs),
        'pfcs': format_number(scope1_pfcs + scope2_loc_pfcs + scope3_pfcs),
        'sf6': format_number(scope1_sf6 + scope2_loc_sf6 + scope3_sf6),
        'nf3': format_number(scope1_nf3 + scope2_loc_nf3 + scope3_nf3),
        'total': format_number(scope1_total + scope2_loc_total + scope3_total),
    }

    # 基于市场总计 = 范围一 + 范围二(市场) + 范围三
    total_mkt_co2 = scope1_co2 + scope2_mar_co2 + scope3_co2
    total_mkt_ch4 = scope1_ch4 + scope2_mar_ch4 + scope3_ch4
    total_mkt_n2o = scope1_n2o + scope2_mar_n2o + scope3_n2o
    total_mkt_hfcs = scope1_hfcs + scope2_mar_hfcs + scope3_hfcs
    total_mkt_pfcs = scope1_pfcs + scope2_mar_pfcs + scope3_pfcs
    total_mkt_sf6 = scope1_sf6 + scope2_mar_sf6 + scope3_sf6
    total_mkt_nf3 = scope1_nf3 + scope2_mar_nf3 + scope3_nf3
    total_mkt_total = scope1_total + scope2_mar_total + scope3_total

    summary_data['total_mkt'] = {
        'co2': format_number(total_mkt_co2),
        'ch4': format_number(total_mkt_ch4),
        'n2o': format_number(total_mkt_n2o),
        'hfcs': format_number(total_mkt_hfcs),
        'pfcs': format_number(total_mkt_pfcs),
        'sf6': format_number(total_mkt_sf6),
        'nf3': format_number(total_mkt_nf3),
        'total': format_number(total_mkt_total),
    }

    # 添加别名：total_mar (与total_mkt相同，向后兼容)
    summary_data['total_mar'] = summary_data['total_mkt']

    # 默认total使用基于位置的数据（兼容模板中使用summary_data.total的情况）
    summary_data['total'] = summary_data['total_loc']

    return summary_data


# ============================================================
# 测试代码
# ============================================================
if __name__ == '__main__':
    import json

    # Mock 测试数据
    mock_raw_data = {
        # 范围一排放数据
        'scope1_co2': 7086080.35,
        'scope1_ch4': 156.23,
        'scope1_n2o': 89.45,
        'scope1_hfcs': 12345.67,
        'scope1_pfcs': 2345.89,
        'scope1_sf6': 456.78,
        'scope1_nf3': 123.45,
        'scope1_total': 7122248.83,

        # 范围二排放数据
        'scope2_co2': 950000.00,
        'scope2_ch4': 0,
        'scope2_n2o': 0,
        'scope2_hfcs': 0,
        'scope2_pfcs': 0,
        'scope2_sf6': 0,
        'scope2_nf3': 0,
        'scope2_total': 950000.00,

        # 范围三总计（备用，实际由cat1-15累加）
        'scope3_co2': 6000000.00,
        'scope3_ch4': 120000.00,
        'scope3_n2o': 45000.00,
        'scope3_hfcs': 89000.00,
        'scope3_pfcs': 12000.00,
        'scope3_sf6': 3400.00,
        'scope3_nf3': 890.00,
        'scope3_total': 6272290.00,
    }

    # 范围三15个类别的明细数据
    scope3_categories_data = {
        'cat1': {'co2': 1500000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 1500000.00},
        'cat2': {'co2': 800000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 800000.00},
        'cat3': {'co2': 1200000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 1200000.00},
        'cat4': {'co2': 250000.00, 'ch4': 15000.00, 'n2o': 5000.00, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 270000.00},
        'cat5': {'co2': 100000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 100000.00},
        'cat6': {'co2': 80000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 80000.00},
        'cat7': {'co2': 50000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 50000.00},
        'cat8': {'co2': 0, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 0},
        'cat9': {'co2': 180000.00, 'ch4': 20000.00, 'n2o': 8000.00, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 208000.00},
        'cat10': {'co2': 0, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 0},
        'cat11': {'co2': 60000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 60000.00},
        'cat12': {'co2': 40000.00, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 40000.00},
        'cat13': {'co2': 0, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 0},
        'cat14': {'co2': 0, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 0},
        'cat15': {'co2': 0, 'ch4': 0, 'n2o': 0, 'hfcs': 0, 'pfcs': 0, 'sf6': 0, 'nf3': 0, 'total': 0},
    }
    mock_raw_data.update(scope3_categories_data)

    # 绿电/绿证数据
    mock_green_power_data = {
        'gec_mwh': 50000.00,        # 绿证GEC电量
        'green_elec_mwh': 30000.00,  # 绿电
        'other_mwh': 5000.00,       # 其他
    }

    # 生成汇总数据
    result = generate_inventory_context(mock_raw_data, mock_green_power_data)

    # 打印结果
    print("=" * 70)
    print("基准年温室气体清单汇总数据")
    print("=" * 70)

    print("\n【范围一排放数据】")
    print(json.dumps(result['scope1'], indent=2, ensure_ascii=False))

    print("\n【范围三排放数据】（前5个类别 + 总计）")
    scope3_preview = {k: result['scope3'][k] for k in ['cat1', 'cat2', 'cat3', 'cat4', 'cat5', 'total']}
    print(json.dumps(scope3_preview, indent=2, ensure_ascii=False))

    print("\n【绿电数据】")
    print(json.dumps(result['green_power'], indent=2, ensure_ascii=False))

    print("\n【范围二排放 - 基于位置】")
    print(json.dumps(result['scope2']['loc'], indent=2, ensure_ascii=False))

    print("\n【范围二排放 - 基于市场】")
    print(json.dumps(result['scope2']['mar'], indent=2, ensure_ascii=False))

    print("\n【企业总计 - 基于位置】")
    print(json.dumps(result['total_loc'], indent=2, ensure_ascii=False))

    print("\n【企业总计 - 基于市场】")
    print(json.dumps(result['total_mar'], indent=2, ensure_ascii=False))

    print("\n" + "=" * 70)
    print("数据验证：")
    print(f"  基于位置总计: {result['total_loc']['total']}")
    print(f"  基于市场总计: {result['total_mar']['total']}")
    print(f"  差异（绿电减排量）: {float(result['total_loc']['total'].replace(',', '')) - float(result['total_mar']['total'].replace(',', '')):.2f}")
    print("=" * 70)
