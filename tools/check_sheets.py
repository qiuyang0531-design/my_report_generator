import openpyxl

# 加载Excel文件
wb = openpyxl.load_workbook('test_data.xlsx')

# 获取所有工作表名称
sheet_names = wb.sheetnames

print("Excel文件中的工作表名称:")
for name in sheet_names:
    print(f"- {name}")

# 关闭工作簿
wb.close()