# -*- coding: utf-8 -*-
"""
彻底删除空类别及其关联的所有内容
包括：标题段落、单位段落、表格
"""

def get_ultimate_clean_function():
    """
    返回终极版的clean_empty_category_tables函数
    确保彻底删除空类别的所有痕迹
    """
    return '''
def clean_empty_category_tables(doc, context):
    """
    彻底删除没有数据的类别（类别8、13-14、15）的所有内容

    删除策略：
    1. 查找空类别的所有相关段落和表格
    2. 删除标题段落
    3. 删除紧邻的"单位：吨CO2e"段落（无论是在标题前面还是后面）
    4. 删除空表格
    5. 扫描整个文档，删除遗留的孤立"单位：吨CO2e"段落
    """
    # 范围三类别数量常量
    TOTAL_SCOPE3_CATEGORIES = 15

    # 类别编号到名称的映射
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
    for i in range(1, TOTAL_SCOPE3_CATEGORIES + 1):
        detail_items = context.get(f'scope3_category{i}', [])
        emission_value = context.get(f'scope_3_category_{i}_emissions', 0)
        has_detail_items = detail_items and len(detail_items) > 0
        has_emissions = emission_value and emission_value > 0

        if not (has_detail_items or has_emissions):
            empty_categories.append(i)

    if not empty_categories:
        print("  所有类别都有数据，无需删除空类别表格")
        return

    print(f"  没有数据的类别: {empty_categories}")

    deleted_count = 0

    for cat_num in empty_categories:
        category_name = category_names.get(cat_num, "")

        # 步骤1：查找并删除所有相关的标题段落
        paragraphs_to_remove = []
        unit_paragraphs_to_remove = []

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()

            # 多种匹配模式
            is_target_category = (
                f'范围三 类别{cat_num}' in text or
                f'范围三类别{cat_num}' in text or
                f'类别{cat_num}' in text or
                category_name in text
            )

            if is_target_category:
                paragraphs_to_remove.append(i)

        # 步骤2：查找与空类别关联的"单位：吨CO2e"段落
        # 检查每个标题段落的前后
        for title_idx in paragraphs_to_remove:
            # 检查前一个段落
            if title_idx > 0:
                prev_para = doc.paragraphs[title_idx - 1]
                if '单位：吨CO2e' in prev_para.text or '单位: 吨CO2e' in prev_para.text:
                    if title_idx - 1 not in unit_paragraphs_to_remove:
                        unit_paragraphs_to_remove.append(title_idx - 1)

            # 检查后一个段落
            if title_idx + 1 < len(doc.paragraphs):
                next_para = doc.paragraphs[title_idx + 1]
                if '单位：吨CO2e' in next_para.text or '单位: 吨CO2e' in next_para.text:
                    if title_idx + 1 not in unit_paragraphs_to_remove:
                        unit_paragraphs_to_remove.append(title_idx + 1)

        # 合并所有要删除的段落
        all_to_remove = sorted(set(paragraphs_to_remove + unit_paragraphs_to_remove), reverse=True)

        # 步骤3：删除所有标记的段落
        for idx in all_to_remove:
            if idx < len(doc.paragraphs):
                para = doc.paragraphs[idx]
                para_element = para._element
                parent = para_element.getparent()
                parent.remove(para_element)
                deleted_count += 1

        # 在删除位置插入空段落保持结构
        if paragraphs_to_remove:
            parent = doc.paragraphs[paragraphs_to_remove[0]]._element.getparent()
            new_para = OxmlElement('w:p')
            parent.insert(0, new_para)

    # 步骤4：删除空类别相关的表格
    tables_to_remove = []
    for table_idx, table in enumerate(doc.tables):
        table_text = ""
        for row in table.rows[:3]:
            for cell in row.cells:
                table_text += cell.text + " "

        # 检查表格是否包含空类别
        for cat_num in empty_categories:
            category_name = category_names.get(cat_num, "")
            if (category_name in table_text or
                f'类别{cat_num}' in table_text):
                if len(table.rows) <= 3:  # 空表格
                    tables_to_remove.append(table_idx)
                    print(f"  标记删除类别{cat_num}的表格: 索引{table_idx}")
                    break

    # 从后往前删除表格
    for table_idx in sorted(tables_to_remove, reverse=True):
        if table_idx < len(doc.tables):
            table = doc.tables[table_idx]
            table_element = table._element
            table_element.getparent().remove(table_element)
            deleted_count += 1

    # 步骤5：最终扫描 - 删除孤立的"单位：吨CO2e"段落
    # 检查所有"单位：吨CO2e"段落，如果它们不属于有数据的类别，则删除
    isolated_units = []
    for i, para in enumerate(doc.paragraphs):
        if '单位：吨CO2e' in para.text or '单位: 吨CO2e' in para.text:
            # 检查前后是否有有效的内容
            has_valid_content = False

            # 检查前5行和后5行
            start = max(0, i - 5)
            end = min(len(doc.paragraphs), i + 6)

            for j in range(start, end):
                if j == i:
                    continue
                text = doc.paragraphs[j].text.strip()
                # 检查是否有有效的类别内容
                for cat_num in range(1, 16):
                    if cat_num not in empty_categories:
                        if f'类别{cat_num}' in text or f'范围三 类别{cat_num}' in text:
                            has_valid_content = True
                            break
                if has_valid_content:
                    break

            if not has_valid_content:
                isolated_units.append(i)

    # 删除孤立的单位段落
    for idx in sorted(isolated_units, reverse=True):
        if idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            para_element = para._element
            para_element.getparent().remove(para_element)
            deleted_count += 1

    print(f"  彻底删除完成，共删除 {deleted_count} 个元素")
    print(f"  包括标题段落、单位段落、表格和孤立单位段落")
'''

# 执行替换
with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

import re
old_pattern = r'def clean_empty_category_tables\(doc, context\):.*?(?=\n\ndef [a-z_]|\nclass [A-Z]|\nif __name__|$)'
new_function = get_ultimate_clean_function()

new_content = re.sub(old_pattern, new_function, content, flags=re.DOTALL)

if new_content != content:
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('[OK] 已更新clean_empty_category_tables函数')
else:
    print('[SKIP] 替换失败')

print('修复完成！')
