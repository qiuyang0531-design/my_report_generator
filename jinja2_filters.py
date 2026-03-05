#!/usr/bin/env python3
"""
Jinja2 自定义过滤器模块

提供数字格式化等展示层过滤器，保持数据层纯净
"""
import jinja2


def format_number(value, decimals=2, with_comma=True):
    """
    格式化数字：添加千分位分隔符，保留小数位（展示层格式化）

    Args:
        value: 数值（float, int 或 str）
        decimals: 小数位数，默认 2
        with_comma: 是否添加千分位分隔符，默认 True

    Returns:
        格式化后的字符串，如 "7,122,248.83"
        无效值返回 "0.00"

    Examples:
        >>> format_number(7122248.834)
        '7,122,248.83'
        >>> format_number(7122248.834, decimals=1)
        '7,122,248.8'
        >>> format_number(7122248.834, with_comma=False)
        '7122248.83'
        >>> format_number("0.00")
        '0.00'
        >>> format_number(None)
        '0.00'
    """
    if value is None:
        return "0.00"

    try:
        float_value = float(value)

        # 处理零值
        if float_value == 0:
            return "0.00"

        # 格式化数字
        if with_comma:
            # 添加千分位分隔符
            return f"{float_value:,.{decimals}f}"
        else:
            # 不添加千分位分隔符
            return f"{float_value:.{decimals}f}"
    except (ValueError, TypeError):
        return "0.00"


def format_emission(value, decimals=2):
    """
    格式化排放量数据（专门用于排放数据）

    与 format_number 类似，但对于零值返回空字符串，
    用于在模板中不显示无数据的排放项

    Args:
        value: 排放量数值
        decimals: 小数位数，默认 2

    Returns:
        格式化后的字符串，零值返回空字符串

    Examples:
        >>> format_emission(7122248.83)
        '7,122,248.83'
        >>> format_emission(0)
        ''
        >>> format_emission(None)
        ''
    """
    if value is None or value == 0:
        return ""

    try:
        float_value = float(value)
        if float_value > 0:
            return f"{float_value:,.{decimals}f}"
        else:
            return ""
    except (ValueError, TypeError):
        return ""


def format_percent(value, decimals=2):
    """
    格式化百分比数据

    Args:
        value: 百分比值（如 0.156 表示 15.6%）
        decimals: 小数位数，默认 2

    Returns:
        格式化后的百分比字符串

    Examples:
        >>> format_percent(0.156)
        '15.60%'
        >>> format_percent(0.156, decimals=1)
        '15.6%'
    """
    try:
        float_value = float(value)
        return f"{float_value * 100:.{decimals}f}%"
    except (ValueError, TypeError):
        return "0.00%"


def format_yes_no(value):
    """
    将布尔值或数值转换为"是/否"文本

    Args:
        value: 布尔值或数值

    Returns:
        "是" 或 "否"

    Examples:
        >>> format_yes_no(True)
        '是'
        >>> format_yes_no(1)
        '是'
        >>> format_yes_no(0)
        '否'
        >>> format_yes_no(False)
        '否'
    """
    try:
        if bool(value):
            return "是"
        else:
            return "否"
    except:
        return "否"


def register_filters_to_template(docx_template):
    """
    将所有自定义过滤器注册到 DocxTemplate 的 Jinja2 环境

    Args:
        docx_template: DocxTemplate 实例

    Returns:
        DocxTemplate 实例（已注册过滤器）

    Usage:
        >>> from docxtpl import DocxTemplate
        >>> from jinja2_filters import register_filters_to_template
        >>>
        >>> template = DocxTemplate("template.docx")
        >>> register_filters_to_template(template)
        >>>
        >>> # 现在可以在模板中使用过滤器了
        >>> # {{ scope_1_emissions|format_number }}
        >>> # {{ scope_2_emissions|format_emission }}
    """
    # 获取或创建 Jinja2 环境
    if not hasattr(docx_template, 'jinja_env') or docx_template.jinja_env is None:
        docx_template.jinja_env = jinja2.Environment()

    # 注册过滤器
    docx_template.jinja_env.filters['format_number'] = format_number
    docx_template.jinja_env.filters['format_emission'] = format_emission
    docx_template.jinja_env.filters['format_percent'] = format_percent
    docx_template.jinja_env.filters['format_yes_no'] = format_yes_no

    return docx_template


# 如果直接运行此模块，提供测试代码
if __name__ == "__main__":
    print("Jinja2 自定义过滤器测试")
    print("=" * 80)

    # 测试 format_number
    print("\n测试 format_number:")
    print(f"  format_number(7122248.834) = {format_number(7122248.834)}")
    print(f"  format_number(0) = {format_number(0)}")
    print(f"  format_number(None) = {format_number(None)}")
    print(f"  format_number(1234.5, with_comma=False) = {format_number(1234.5, with_comma=False)}")

    # 测试 format_emission
    print("\n测试 format_emission:")
    print(f"  format_emission(7122248.83) = {format_emission(7122248.83)}")
    print(f"  format_emission(0) = '{format_emission(0)}'")  # 应该是空字符串
    print(f"  format_emission(None) = '{format_emission(None)}'")

    # 测试 format_percent
    print("\n测试 format_percent:")
    print(f"  format_percent(0.156) = {format_percent(0.156)}")
    print(f"  format_percent(0.156, decimals=1) = {format_percent(0.156, decimals=1)}")

    # 测试 format_yes_no
    print("\n测试 format_yes_no:")
    print(f"  format_yes_no(True) = {format_yes_no(True)}")
    print(f"  format_yes_no(False) = {format_yes_no(False)}")
    print(f"  format_yes_no(1) = {format_yes_no(1)}")
    print(f"  format_yes_no(0) = {format_yes_no(0)}")

    print("\n" + "=" * 80)
    print("✓ 所有过滤器测试完成")
