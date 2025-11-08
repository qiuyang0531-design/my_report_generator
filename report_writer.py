from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn # 关键：用于处理中文字体
from docx.enum.text import WD_ALIGN_PARAGRAPH 

class WordReportWriter: 
    def __init__(self): 
        """
        初始化时，创建一个空白文档，并立即设置好样式。
        """
        self.doc = Document() 
        print("创建新 Word 文档") 
        self._setup_styles()
        self._setup_page_margins() 

    def _setup_styles(self): 
        """
        "样式与内容分离"思想的核心。
        在这里定义好所有"规则"。
        """
        print("开始设置 Word 样式...") 
        
        # 1. 设置默认（正文）样式 
        style = self.doc.styles['Normal'] 
        style.font.name = '宋体' 
        # 关键一步：设置东亚字体（处理中文字符） 
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体') 
        style.font.size = Pt(10.5)
        # 设置段落间距
        style.paragraph_format.space_after = Pt(0)  # 段后0点
        style.paragraph_format.line_spacing = 1.0  # 单倍行距

        # 2. 设置一级标题样式 
        style = self.doc.styles['Heading 1'] 
        style.font.name = '黑体' 
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体') 
        style.font.size = Pt(16)
        style.paragraph_format.space_after = Pt(12)  # 段后12点
        
        print("样式设置完成。")
        
    def _setup_page_margins(self):
        """
        设置页面边距，匹配模板文档格式。
        """
        print("设置页面边距...")
        sections = self.doc.sections
        for section in sections:
            # 设置标准边距（1英寸 = 2.54厘米）
            section.top_margin = Inches(1.25)  # 上边距
            section.bottom_margin = Inches(1.0)  # 下边距
            section.left_margin = Inches(1.25)  # 左边距
            section.right_margin = Inches(1.25)  # 右边距
        print("页面边距设置完成。") 

    def add_title_page(self, company_name, report_year):
        """
        添加封面页，完全参照模板1的格式
        """
        import os
        from datetime import datetime
        from docx.shared import Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # 添加封面图片（如果存在）
        script_dir = os.path.dirname(os.path.abspath(__file__))
        cover_image_path = os.path.join(script_dir, "封面.png")

        print(f"尝试查找封面图片: {cover_image_path}")
        print(f"封面图片是否存在: {os.path.exists(cover_image_path)}")

        # 1. 添加几个空行作为顶部留白
        for _ in range(2):
            self.doc.add_paragraph()

        # 2. 添加封面图片（居中）
        if os.path.exists(cover_image_path):
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run()
            # 添加图片，设置适当的宽度（5.5英寸）
            run.add_picture(cover_image_path, width=Inches(5.5))
            print("已添加封面图片")

            # 在图片后添加一些空行
            for _ in range(2):
                self.doc.add_paragraph()
        else:
            # 如果没有图片，添加更多空行作为留白
            for _ in range(4):
                self.doc.add_paragraph()

        # 3. 公司名称 - 26pt，加粗，居中
        company_paragraph = self.doc.add_paragraph()
        company_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        company_run = company_paragraph.add_run(company_name)
        company_run.font.size = Pt(26)  # 模板中的26pt
        company_run.font.bold = True
        company_run.font.name = '黑体'
        company_run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

        # 4. 报告标题 - 26pt，加粗，居中
        title_paragraph = self.doc.add_paragraph()
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_paragraph.add_run("温室气体排放盘查报告")
        title_run.font.size = Pt(26)  # 模板中的26pt
        title_run.font.bold = True
        title_run.font.name = '黑体'
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

        # 添加空行
        for _ in range(3):
            self.doc.add_paragraph()

        # 5. 统计期间 - 14pt，左对齐
        period_paragraph = self.doc.add_paragraph()
        period_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        period_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        period_run = period_paragraph.add_run(f"统计期间：{report_year}年1月1日~ {report_year}年12月31日")
        period_run.font.size = Pt(14)  # 模板中的14pt
        period_run.font.name = '宋体'
        period_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 6. 版本号 - 14pt，左对齐
        version_paragraph = self.doc.add_paragraph()
        version_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        version_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        version_run = version_paragraph.add_run("版本号：A/1")
        version_run.font.size = Pt(14)
        version_run.font.name = '宋体'
        version_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 7. 文件编号 - 14pt，左对齐
        file_number_paragraph = self.doc.add_paragraph()
        file_number_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        file_number_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        # 生成基于当前日期的文件编号
        current_date = datetime.now()
        file_number = f"DY-GHG-{current_date.strftime('%Y')}-01"
        file_number_run = file_number_paragraph.add_run(f"文件编号：{file_number}")
        file_number_run.font.size = Pt(14)
        file_number_run.font.name = '宋体'
        file_number_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 8. 编制日期 - 14pt，左对齐
        create_date_paragraph = self.doc.add_paragraph()
        create_date_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        create_date_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        create_date_run = create_date_paragraph.add_run(f"编制日期：{current_date.strftime('%Y年%m月%d日')}")
        create_date_run.font.size = Pt(14)
        create_date_run.font.name = '宋体'
        create_date_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 9. 修订日期 - 14pt，左对齐
        revise_date_paragraph = self.doc.add_paragraph()
        revise_date_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        revise_date_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        revise_date_run = revise_date_paragraph.add_run(f"修订日期：{current_date.strftime('%Y年%m月%d日')}")
        revise_date_run.font.size = Pt(14)
        revise_date_run.font.name = '宋体'
        revise_date_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 添加大量空行将签名区域推到页面底部
        for _ in range(8):
            self.doc.add_paragraph()

        # 10. 签名区域标题 - 14pt，居中
        signature_header_paragraph = self.doc.add_paragraph()
        signature_header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        signature_header_run = signature_header_paragraph.add_run("审核人：                          编制人：                          批准人：")
        signature_header_run.font.size = Pt(14)
        signature_header_run.font.name = '宋体'
        signature_header_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        # 11. 底部信息 - 统一设置对齐和段落格式

        # 审核信息
        review_paragraph = self.doc.add_paragraph()
        review_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        review_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        review_run = review_paragraph.add_run("审核")
        review_run.font.size = Pt(14)
        review_run.font.bold = True
        review_run.font.name = '宋体'
        review_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        review_paragraph.paragraph_format.line_spacing = 1.5

        # 组织
        org_paragraph = self.doc.add_paragraph()
        org_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        org_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        org_run = org_paragraph.add_run("审核组织")
        org_run.font.size = Pt(12)  # 默认字体大小
        org_run.font.name = '宋体'
        org_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        org_paragraph.paragraph_format.line_spacing = 1.5

        # 公司名称
        company_info_paragraph = self.doc.add_paragraph()
        company_info_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        company_info_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        company_info_run = company_info_paragraph.add_run(f"公司名称：{company_name}")
        company_info_run.font.size = Pt(12)
        company_info_run.font.bold = True  # 标签加粗
        company_info_run.font.name = '宋体'
        company_info_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        company_info_paragraph.paragraph_format.line_spacing = 1.5

        # 统一社会信用代码（使用示例代码）
        credit_code_paragraph = self.doc.add_paragraph()
        credit_code_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        credit_code_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        credit_code_run = credit_code_paragraph.add_run("统一社会信用代码：91420000798750168P")
        credit_code_run.font.size = Pt(12)
        credit_code_run.font.bold = True  # 标签加粗
        credit_code_run.font.name = '宋体'
        credit_code_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        credit_code_paragraph.paragraph_format.line_spacing = 1.5

        # 编制人
        preparer_paragraph = self.doc.add_paragraph()
        preparer_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        preparer_paragraph.paragraph_format.space_after = Pt(0)  # 段后不留空
        preparer_run = preparer_paragraph.add_run("编制人：系统自动生成")
        preparer_run.font.size = Pt(12)
        preparer_run.font.bold = True  # 标签加粗
        preparer_run.font.name = '宋体'
        preparer_run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        preparer_paragraph.paragraph_format.line_spacing = 1.5

        print(f"已添加模板格式的封面: {company_name}")

        # 添加分页符
        self.doc.add_page_break() 

    def add_emission_table(self, data): 
        """
        接收数据字典，并动态生成一个表格。
        """
        print("开始生成排放汇总表...") 
        self.doc.add_heading("1. 排放数据汇总", level=1) 
        
        # 准备要填入表格的数据 
        # 支持data_reader.py返回的完整数据结构
        table_data = [ 
            ("总排放量(基于位置)(tCO2e)", data.get('total_emission_location', '未找到')), 
            ("总排放量(基于市场)(tCO2e)", data.get('total_emission_market', '未找到')), 
            ("范围一(tCO2e)", data.get('scope_1', '未找到')), 
            ("范围二(基于位置)(tCO2e)", data.get('scope_2_location', '未找到')), 
            ("范围二(基于市场)(tCO2e)", data.get('scope_2_market', '未找到')), 
            ("范围三(tCO2e)", data.get('scope_3', '未找到')) 
        ] 

        # 创建一个表格（1行表头 + 6行数据 = 7行, 2列） 
        table = self.doc.add_table(rows=1, cols=2) 
        table.style = 'Table Grid' # 带网格线的表格样式 

        # 1. 填充表头 
        hdr_cells = table.rows[0].cells 
        hdr_cells[0].text = '排放项目' 
        hdr_cells[1].text = '排放量'
        
        # 2. 设置表头样式
        hdr_cells[0].paragraphs[0].runs[0].bold = True
        hdr_cells[1].paragraphs[0].runs[0].bold = True

        # 3. 填充数据行 
        for item, value in table_data: 
            row_cells = table.add_row().cells 
            row_cells[0].text = item 
            # 确保所有值都是字符串 
            row_cells[1].text = str(value) 

        print("表格生成完毕。") 

    def save(self, output_path): 
        """
        保存最终生成的 Word 文档，添加错误处理以处理文件锁定等情况。
        """
        try:
            # 尝试直接保存
            self.doc.save(output_path)
            print(f"文档已成功保存到: {output_path}")
            return True
        except PermissionError:
            # 如果文件被锁定，尝试使用时间戳重命名保存
            import os
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name, ext = os.path.splitext(output_path)
            new_output_path = f"{base_name}_{timestamp}{ext}"
            self.doc.save(new_output_path)
            print(f"原文件被锁定，已保存为: {new_output_path}")
            return True
        except Exception as e:
            print(f"保存文件失败: {e}")
            return False

# (你可以在文件末尾添加测试代码) 
if __name__ == "__main__": 
    writer = WordReportWriter() 
    test_data = { 
        'total_emission': 1000, 
        'scope_1': 600, 
        'scope_2': 400 
    } 
    writer.add_title_page("我的测试公司", 2024) 
    writer.add_emission_table(test_data) 
    writer.save("test_report.docx") 
    print("--- 测试 report_writer.py 完成 ---")