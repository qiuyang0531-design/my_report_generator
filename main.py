# main.py
from data_reader import ExcelDataReader
from docxtpl import DocxTemplate
import os


def format_number(value, decimals=2):
    """
    格式化数字：保留指定小数位数（展示层格式化）

    Args:
        value: 数值（float 或 int）
        decimals: 小数位数，默认 2

    Returns:
        格式化后的字符串
    """
    try:
        return f"{float(value):.{decimals}f}"
    except (ValueError, TypeError):
        return "0.00"


def prepare_context_with_formatting(context):
    """
    为 context 添加格式化后的数值版本（展示层处理）
    保持原始数据不变，添加 _formatted 后缀的格式化版本
    对于排放量为0的类别，添加说明文字

    Args:
        context: 原始数据字典

    Returns:
        包含格式化字段的新字典
    """
    formatted_context = context.copy()

    # 计算范围三类别排放总量（用于表格汇总）
    scope3_total_sum = 0
    for i in range(1, 16):
        cat_emission = context.get(f'scope_3_category_{i}_emissions', 0)
        if cat_emission and cat_emission > 0:
            scope3_total_sum += cat_emission

    formatted_context['scope3_categories_total_sum'] = scope3_total_sum
    formatted_context['scope3_categories_total_sum_formatted'] = format_number(scope3_total_sum)

    # 计算每个类别的详细项排放量总和（用于各类别表格汇总）
    # 每个类别计算所有列的总和
    emission_columns = [
        'total_green_house_gas_emissions',
        'CO2_emissions',
        'CH4_emissions',
        'N2O_emissions',
        'HFCs_emissions',
        'PFCs_emissions',
        'SFs_emissions',
        'NF3_emissions'
    ]

    # 列名映射：数据源使用SF6，模板使用SFs
    column_key_mapping = {
        'SFs_emissions': 'SF6_emissions'  # 从数据源读取SF6，但输出为SFs
    }

    for cat_num in range(1, 16):
        category_var = f'scope3_category{cat_num}'
        if category_var in context and context[category_var]:
            # 初始化各列的总和
            column_sums = {col: 0.0 for col in emission_columns}

            # 遍历该类别的所有详细项，累加各列值
            for item in context[category_var]:
                for col in emission_columns:
                    # 根据映射获取实际的数据键名
                    data_key = column_key_mapping.get(col, col)
                    emission_str = item.get(data_key, '0')
                    # 去除逗号和空格，转换为浮点数
                    emission_str = emission_str.replace(',', '').replace(' ', '')
                    try:
                        emission_value = float(emission_str)
                        column_sums[col] += emission_value
                    except (ValueError, TypeError):
                        pass

            # 为每个列创建格式化后的总和变量
            for col in emission_columns:
                sum_value = column_sums[col]
                formatted_context[f'scope3_category{cat_num}_{col}_sum'] = sum_value
                formatted_context[f'scope3_category{cat_num}_{col}_sum_formatted'] = format_number(sum_value)

    # 需要格式化的数值字段列表
    number_fields = [
        'scope_1_emissions',
        'scope_2_location_based_emissions',
        'scope_2_market_based_emissions',
        'scope_3_emissions',
        'total_emission_location',
        'total_emission_market',
    ]

    # 范围三类别字段
    for i in range(1, 16):
        number_fields.append(f'scope_3_category_{i}_emissions')

    # 别名也需要格式化
    alias_fields = [
        'scope_1',
        'scope_2_location',
        'scope_2_market',
        'scope_3',
    ]

    all_number_fields = number_fields + alias_fields

    # 为每个数值字段创建格式化版本
    for field in all_number_fields:
        if field in context and context[field] is not None:
            formatted_context[f'{field}_formatted'] = format_number(context[field])
        else:
            formatted_context[f'{field}_formatted'] = format_number(0)

    # 为范围三类别添加 display 字段（处理0值显示说明文字）
    # 有数据的类别：显示格式化数字
    for i in range(1, 16):
        field = f'scope_3_category_{i}_emissions'
        value = context.get(field, 0)
        if value and value > 0:
            formatted_context[f'{field}_display'] = format_number(value)
        else:
            # 无数据的类别：设置为空，不单独显示
            formatted_context[f'{field}_display'] = ""

    # 收集所有无数据的范围三类别，生成汇总说明
    categories_not_in_scope = []
    for i in range(1, 16):
        field = f'scope_3_category_{i}_emissions'
        value = context.get(field, 0)
        if not value or value <= 0:
            categories_not_in_scope.append(i)

    # 生成汇总说明文字
    if categories_not_in_scope:
        category_str = "、".join(f"类别{cat}" for cat in categories_not_in_scope)
        formatted_context['scope_3_categories_not_in_scope_summary'] = (
            f"范围三{category_str}产生的排放为0，排放量不在本次盘查范围内，本周期内不进行量化。"
        )
    else:
        formatted_context['scope_3_categories_not_in_scope_summary'] = ""

    return formatted_context


def generate_report_from_xlsx(
    xlsx_path="test_data.xlsx",
    template_path="template.docx",
    output_path="carbon_report.docx"
):
    """
    使用 template.docx 作为模板，从 xlsx 文件动态读取数据生成报告
    数字格式化在展示层通过 Jinja2 过滤器处理，数据层保持原始数字类型

    Args:
        xlsx_path: Excel 数据文件路径
        template_path: Word 模板文件路径
        output_path: 输出报告路径（默认: carbon_report.docx）
    """
    print("=" * 50)
    print("开始生成碳盘查报告（纯xlsx，动态读取）")
    print("=" * 50)

    # 1. 使用 data_reader 的动态读取方法
    print(f"\n[步骤1] 从 {xlsx_path} 动态提取数据...")
    reader = ExcelDataReader(xlsx_path)
    context = reader.extract_data_from_xlsx_dynamic()

    # 打印提取的关键数据
    print("\n提取的关键数据:")
    print(f"  公司名称: {context.get('company_name')}")
    print(f"  范围一排放: {context.get('scope_1_emissions')}")
    print(f"  范围二排放（基于位置）: {context.get('scope_2_location_based_emissions')}")
    print(f"  范围二排放（基于市场）: {context.get('scope_2_market_based_emissions')}")
    print(f"  范围三排放: {context.get('scope_3_emissions')}")

    print("\n范围三分类排放量:")
    for i in range(1, 16):  # 范围三共15个类别
        val = context.get(f'scope_3_category_{i}_emissions')
        if val and val > 0:
            print(f"  类别{i}: {val}")

    # 2. 加载模板
    print(f"\n[步骤2] 加载模板: {template_path}")
    template = DocxTemplate(template_path)

    # 3. 准备展示层数据（格式化数字，保持数据层纯净）
    print("\n[步骤3] 准备展示层数据（格式化数字）...")
    render_context = prepare_context_with_formatting(context)

    # 4. 渲染模板
    print("[步骤4] 渲染模板...")
    template.render(render_context)

    # 5. 保存报告
    print(f"\n[步骤5] 保存报告到: {output_path}")
    template.save(output_path)

    # 6. 统一公司简介和经营范围的段落格式
    print(f"\n[步骤6] 统一段落格式...")
    from docx import Document
    from docx.shared import Pt, Inches, Cm

    doc = Document(output_path)
    company_name = context.get('company_name', '')

    # 设置统一的首行缩进：4个空格 ≈ 2个中文字符 ≈ 0.4厘米
    first_line_indent = Cm(0.4)  # 约2个中文字符的宽度

    # 遍历所有段落，精确处理公司简介和经营范围
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # 检查是否是公司简介段落（以公司名开头或包含公司简介特征）
        if company_name in text[:50] or '楚能新能源' in text[:50]:
            if len(text) > 100:  # 确保是长文本内容而非标题
                # 设置统一的首行缩进，左缩进为0
                para.paragraph_format.first_line_indent = first_line_indent
                para.paragraph_format.left_indent = 0
                print(f"  已处理公司简介段落（长度: {len(text)} 字符）")

        # 检查是否是经营范围段落（包含经营范围特征）
        elif '经营范围' in text and len(text) > 100:
            # 设置统一的首行缩进，左缩进为0
            para.paragraph_format.first_line_indent = first_line_indent
            para.paragraph_format.left_indent = 0
            print(f"  已处理经营范围段落（长度: {len(text)} 字符）")

    # 7. 设置表2的列宽相等
    print(f"\n[步骤7] 设置表2列宽...")
    # 找到表2（范围二、三间接排放源表格）
    if len(doc.tables) >= 2:
        table2 = doc.tables[1]  # 第二个表格是表2
        # 设置所有列宽相等（使用 Inches 单位）
        equal_width = Inches(2.0)  # 每列2英寸
        for row in table2.rows:
            for cell in row.cells:
                # 设置单元格宽度
                cell.width = equal_width
        print("  表2列宽已设置为相等")

    doc.save(output_path)
    print("段落格式统一完成")

    # 8. 删除没有数据的类别表格（类别8、13、14、15）
    print(f"\n[步骤8] 删除没有数据的类别表格...")
    clean_empty_category_tables(doc, context)

    doc.save(output_path)
    print("空类别表格清理完成")

    print("\n" + "=" * 50)
    print(f"报告生成成功: {output_path}")
    print("=" * 50)

    return output_path


def find_table_by_content(doc, search_keywords):
    """
    根据表格内容动态查找表格索引

    Args:
        doc: Word文档对象
        search_keywords: 要搜索的关键词列表（表格中包含任一关键词即匹配）

    Returns:
        匹配的表格索引，如果未找到返回None
    """
    for idx, table in enumerate(doc.tables):
        # 获取表格中所有文本
        table_text = ""
        for row in table.rows:
            for cell in row.cells:
                table_text += cell.text + " "

        # 检查是否包含任一关键词
        for keyword in search_keywords:
            if keyword in table_text:
                return idx

    return None


def find_summary_table(doc):
    """
    查找范围三排放汇总表格

    Returns:
        汇总表格的索引，如果未找到返回None
    """
    for idx, table in enumerate(doc.tables):
        # 汇总表格的特征：包含多个"范围三 类别X"和"排放量(tCO2e)"等
        table_text = ""
        for row in table.rows:
            for cell in row.cells:
                table_text += cell.text + " "

        # 检查是否是汇总表格（包含至少5个类别，且有"排放量"列）
        category_count = table_text.count("范围三 类别")
        if category_count >= 5 and "排放量" in table_text:
            return idx

    return None


def clean_empty_category_tables(doc, context):
    """
    删除没有数据的类别表格和段落（类别8、13、14、15），并重新编号
    使用动态表格查找，不依赖硬编码的表格索引
    """
    # 范围三类别数量常量
    TOTAL_SCOPE3_CATEGORIES = 15

    # 类别编号到名称的映射（用于动态查找表格）
    category_names = {
        1: "购买的商品和服务",
        2: "资本商品",
        3: "燃料和能源相关活动",
        4: "上游运输和配送",
        5: "运营中产生的废弃物",
        6: "员工商务旅行",
        7: "员工通勤",
        8: "上游租赁资产",
        9: "下游运输和配送",
        10: "销售产品的加工",
        11: "销售产品的使用",
        12: "寿命终结处理",
        13: "下游租赁资产",
        14: "特许经营",
        15: "投资"
    }

    # 检查哪些类别没有数据
    empty_categories = []
    has_data_categories = []  # 有数据的类别
    category_map = {}
    for i in range(1, TOTAL_SCOPE3_CATEGORIES + 1):
        category_map[f'scope3_category{i}'] = i

    for var_name, cat_num in category_map.items():
        if context.get(var_name) and len(context.get(var_name, [])) > 0:
            has_data_categories.append(cat_num)
        else:
            empty_categories.append(cat_num)

    if not empty_categories:
        print("  所有类别都有数据，无需删除")
        return

    print(f"  没有数据的类别: {empty_categories}")
    print(f"  有数据的类别: {has_data_categories}")

    # 1. 删除空类别的段落
    print("  正在删除空类别的段落...")
    paragraphs_to_remove = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        # 检查是否是空类别的段落
        for cat_num in empty_categories:
            if f'范围三类别{cat_num}' in text or f'类别{cat_num}' in text:
                # 标记要删除的段落
                if i not in paragraphs_to_remove:
                    paragraphs_to_remove.append(i)
                # 同时删除下一段（通常是"单位：吨CO2e"）
                if i + 1 < len(doc.paragraphs) and i + 1 not in paragraphs_to_remove:
                    paragraphs_to_remove.append(i + 1)
                break

    # 从大到小删除段落
    for idx in sorted(paragraphs_to_remove, reverse=True):
        if idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            para_element = para._element
            para_element.getparent().remove(para_element)

    print(f"  已删除 {len(paragraphs_to_remove)} 个段落")

    # 2. 动态查找并删除空类别的表格
    print("  正在查找并删除空类别的表格...")
    tables_to_remove = []

    for cat_num in empty_categories:
        # 根据类别编号和名称动态查找表格
        category_name = category_names.get(cat_num, "")
        search_keywords = [
            f'范围三 类别{cat_num}',
            f'范围三类别{cat_num}',
            f'类别{cat_num}',
            category_name
        ]

        table_idx = find_table_by_content(doc, search_keywords)
        if table_idx is not None and table_idx not in tables_to_remove:
            tables_to_remove.append(table_idx)
            print(f"  找到类别{cat_num}的表格: 索引{table_idx}")

    # 从大到小删除表格（避免索引变化）
    for idx in sorted(tables_to_remove, reverse=True):
        if idx < len(doc.tables):
            table = doc.tables[idx]
            table_element = table._element
            table_element.getparent().remove(table_element)
            print(f"  已删除表格 {idx}")

    # 3. 动态查找汇总表格并删除对应行
    summary_table_idx = find_summary_table(doc)
    if summary_table_idx is not None:
        print(f"  找到汇总表格: 索引{summary_table_idx}")
        summary_table = doc.tables[summary_table_idx]
        rows_to_remove = []

        for row_idx, row in enumerate(summary_table.rows):
            row_text = row.cells[0].text.strip()
            # 检查是否是空类别的行
            is_empty_category = False
            for cat_num in empty_categories:
                if f'范围三 类别{cat_num}' in row_text or f'类别{cat_num}' in row_text:
                    is_empty_category = True
                    break

            if is_empty_category:
                # 删除该行和下一行（标题行和数据行）
                if row_idx < len(summary_table.rows):
                    rows_to_remove.append(row_idx)
                if row_idx + 1 < len(summary_table.rows):
                    rows_to_remove.append(row_idx + 1)

        # 去重并从大到小删除行
        for row_idx in sorted(set(rows_to_remove), reverse=True):
            if row_idx < len(summary_table.rows):
                row = summary_table.rows[row_idx]
                row._element.getparent().remove(row._element)
        print(f"  已删除汇总表格中的 {len(set(rows_to_remove))} 行")
    else:
        print("  警告: 未找到汇总表格")

    # 注意：不再重新编号类别，保持原有编号不变
    # 这样类别编号与数据源保持一致，避免数据错位
    print(f"  清理完成，类别编号保持不变（允许有空缺）")


if __name__ == "__main__":
    import sys

    # 检查命令行参数
    if len(sys.argv) > 1 and sys.argv[1] == '--generate':
        # 生成报告模式
        xlsx_path = sys.argv[2] if len(sys.argv) > 2 else "test_data.xlsx"
        output_path = sys.argv[3] if len(sys.argv) > 3 else "carbon_report.docx"
        generate_report_from_xlsx(xlsx_path=xlsx_path, output_path=output_path)
    else:
        # 默认执行生成报告
        print("使用 'python main.py --generate' 生成报告")
        generate_report_from_xlsx()