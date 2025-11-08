from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

def analyze_template():
    """
    分析模板1.docx的格式和结构
    """
    template_path = "模板1.docx"

    if not os.path.exists(template_path):
        print(f"模板文件 {template_path} 不存在")
        return

    try:
        doc = Document(template_path)
        print(f"===== 分析 {template_path} =====")
        print(f"文档总段落数: {len(doc.paragraphs)}")
        print(f"文档总表格数: {len(doc.tables)}")
        print(f"文档总节数: {len(doc.sections)}")

        # 分析页面设置
        if doc.sections:
            section = doc.sections[0]
            print(f"\n页面设置:")
            print(f"  页面宽度: {section.page_width.inches:.2f} 英寸")
            print(f"  页面高度: {section.page_height.inches:.2f} 英寸")
            print(f"  页边距上: {section.top_margin.inches:.2f} 英寸")
            print(f"  页边距下: {section.bottom_margin.inches:.2f} 英寸")
            print(f"  页边距左: {section.left_margin.inches:.2f} 英寸")
            print(f"  页边距右: {section.right_margin.inches:.2f} 英寸")

        print(f"\n===== 段落分析 =====")
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():  # 只显示非空段落
                print(f"段落 {i+1}: '{paragraph.text[:50]}...'" if len(paragraph.text) > 50 else f"段落 {i+1}: '{paragraph.text}'")

                # 分析字体
                if paragraph.runs:
                    for j, run in enumerate(paragraph.runs):
                        if run.text.strip():
                            print(f"  运行 {j+1}: '{run.text[:30]}...'" if len(run.text) > 30 else f"  运行 {j+1}: '{run.text}'")
                            print(f"    字体: {run.font.name}")
                            print(f"    字号: {run.font.size}")
                            print(f"    加粗: {run.font.bold}")
                            print(f"    对齐: {paragraph.alignment}")

                print()

                # 只分析前10个段落避免过长
                if i >= 9:
                    print(f"... (还有 {len(doc.paragraphs) - 10} 个段落)")
                    break

        # 检查是否有图片
        print(f"\n===== 图片分析 =====")
        image_count = 0
        for i, paragraph in enumerate(doc.paragraphs):
            for run in paragraph.runs:
                if hasattr(run.element, './/pic:pic') or run.element.xpath('.//pic:pic'):
                    image_count += 1
                    print(f"发现图片在段落 {i+1}")

        print(f"总共发现 {image_count} 张图片")

    except Exception as e:
        print(f"分析模板时出错: {e}")

if __name__ == "__main__":
    analyze_template()