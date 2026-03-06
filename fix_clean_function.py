# -*- coding: utf-8 -*-
"""
重写clean_empty_category_tables函数，彻底清理空类别
包括：删除表格及其上方的"单位：吨CO2e"段落和类别标题段落
"""

def get_new_clean_empty_category_tables():
    """
    返回新的clean_empty_category_tables函数代码
    """
    return '''
def clean_empty_category_tables(doc, context):
    """
    彻底删除没有数据的类别表格、标题段落和单位段落（类别8、13、14、15）

    删除策略：
    1. 查找空类别的标题段落（如"范围三 类别8 上游租赁资产"）
    2. 删除标题段落及其紧邻的"单位：吨CO2e"段落
    3. 删除紧随其后的空表格
    4. 在原位置保留一个空段落，防止XML结构塌陷
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

    # 彻底删除空类别
    deleted_count = 0

    for cat_num in empty_categories:
        category_name = category_names.get(cat_num, "")

        # 步骤1：查找并删除类别标题段落和单位段落
        title_deleted = False
        unit_deleted = False

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()

            # 检查是否是目标类别的标题段落
            is_target_category = (
                f'范围三 类别{cat_num}' in text or
                f'范围三类别{cat_num}' in text or
                (f'类别{cat_num}' in text and category_name in text) or
                text == category_name
            )

            if is_target_category:
                # 找到标题段落，现在需要删除：
                # 1. 标题段落本身
                # 2. 可能紧随其后的"单位：吨CO2e"段落
                # 3. 可能紧随其后的空表格

                para_element = para._element
                parent = para_element.getparent()
                index = parent.index(para_element)

                # 删除标题段落
                parent.remove(para_element)
                print(f"  已删除类别{cat_num}的标题段落")
                deleted_count += 1
                title_deleted = True

                # 检查并删除紧随其后的单位段落
                # 由于删除了标题段落，原来的index+1位置现在是下一个元素
                if index < len(parent):
                    next_element = parent[index]
                    if next_element.tag.endswith('}p'):  # 是段落
                        try:
                            next_para = type(para)(next_element, doc)
                            if '单位：吨CO2e' in next_para.text or '单位: 吨CO2e' in next_para.text:
                                parent.remove(next_element)
                                print(f"  已删除类别{cat_num}的单位段落")
                                deleted_count += 1
                                unit_deleted = True
                        except:
                            pass

                # 在原位置插入一个空段落，防止XML结构塌陷
                new_para = OxmlElement('w:p')
                parent.insert(index, new_para)

                break

        # 步骤2：查找并删除空表格
        # 遍历所有表格，查找匹配的空表格
        tables_to_remove = []
        for table_idx, table in enumerate(doc.tables):
            table_text = ""
            for row in table.rows[:2]:  # 只检查前2行
                for cell in row.cells:
                    table_text += cell.text + " "

            # 检查表格是否包含目标类别名称
            if (category_name in table_text or
                f'类别{cat_num}' in table_text or
                f'Category {cat_num}' in table_text):

                # 检查表格是否为空（只有表头或数据很少）
                # 如果行数<=3，认为是空表格
                if len(table.rows) <= 3:
                    tables_to_remove.append(table_idx)
                    print(f"  标记删除类别{cat_num}的表格: 索引{table_idx}")

        # 从后往前删除表格（避免索引变化）
        for table_idx in sorted(tables_to_remove, reverse=True):
            if table_idx < len(doc.tables):
                table = doc.tables[table_idx]
                table_element = table._element
                table_element.getparent().remove(table_element)
                deleted_count += 1
                print(f"  已删除类别{cat_num}的表格")

    print(f"  彻底删除完成，共删除 {deleted_count} 个元素")
    print(f"  XML结构保持完整，空类别已完全清除")
'''

# 执行替换
with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 找到旧的clean_empty_category_tables函数
import re
old_pattern = r'def clean_empty_category_tables\(doc, context\):.*?(?=\n\ndef [a-z_]|\nclass [A-Z]|\nif __name__|$)'
new_function = get_new_clean_empty_category_tables()

# 替换
new_content = re.sub(old_pattern, new_function, content, flags=re.DOTALL)

if new_content != content:
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('[OK] 已重写clean_empty_category_tables函数')
else:
    print('[SKIP] 未找到clean_empty_category_tables函数')

print('修复完成！')
