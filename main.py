from data_reader import ExcelDataReader
from report_writer import WordReportWriter
import os
import traceback

# 定义常量，方便维护
INPUT_EXCEL_PATH = 'test_data.xlsx'
OUTPUT_REPORT_PATH = '碳盘查报告.docx'

def main_workflow():
    """
    执行报告生成的主工作流。
    """
    print("===== 碳盘查报告生成系统 =====")
    print("--- [开始] 报告生成工作流 ---")
    
    # 1. 健壮性检查：确保输入文件存在
    if not os.path.exists(INPUT_EXCEL_PATH):
        print(f"错误：找不到输入文件 '{INPUT_EXCEL_PATH}'，程序终止。")
        return
    
    try:
        # 2. "封装"：实例化"读取专家"
        print(f"正在从 '{INPUT_EXCEL_PATH}' 读取数据...")
        reader = ExcelDataReader(INPUT_EXCEL_PATH)
        
        # 3. "抽象"：获取干净的数据字典
        data = reader.extract_data()
        
        # 健壮性检查：如果数据提取失败，则停止
        if not data or not data.get('company_name'):
            print("错误：无法从 Excel 提取到有效数据，程序终止。")
            return
        
        # 4. 数据验证反馈
        print("数据读取完成！")
        print(f"公司名称: {data.get('company_name', '未找到')}")
        print(f"报告年份: {data.get('report_year', '未找到')}")
        print(f"总排放量(基于位置): {data.get('total_emission_location', '未找到')} tCO2e")
        print(f"总排放量(基于市场): {data.get('total_emission_market', '未找到')} tCO2e")
        
        # 5. "封装"：实例化"撰写专家"
        print(f"\n正在生成Word报告...")
        writer = WordReportWriter()
        
        # 6. "指挥"：命令"撰写专家"工作
        # 注意：我们使用 .get() 方法来安全地获取值，即使某个键不存在也不会报错
        writer.add_title_page(
            company_name=data.get('company_name', '未知公司'),
            report_year=data.get('report_year', '未知年份')
        )
        
        writer.add_emission_table(data)
        
        # (未来你可以在这里添加更多章节...)
        # writer.add_introduction(data)
        # writer.add_analysis(data)
        
        # 7. "收尾"：保存最终产物
        writer.save(OUTPUT_REPORT_PATH)
        
        print("\n[SUCCESS] 报告生成完成！")
        print(f"报告文件: {os.path.abspath(OUTPUT_REPORT_PATH)}")
        print("--- [完成] 报告生成工作流 ---")
        
    except Exception as e:
        print(f"[ERROR] 处理过程中发生错误: {str(e)}")
        traceback.print_exc()

# Python 的标准入口
if __name__ == "__main__":
    main_workflow()
    print("\n===== 程序执行结束 =====")