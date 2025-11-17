# web_api.py
import os
import tempfile
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

# 导入你作业一的"专家"
from data_reader import ExcelDataReader
from report_writer import WordReportWriter
# 导入你刚写的"AI 专家"
from ai_service import AIService

# 初始化 Flask 应用
app = Flask(__name__, template_folder='templates')

# 初始化我们的"专家"
# 我们在程序启动时就初始化好，而不是每次请求都初始化
ai_service = AIService()

@app.route("/")
def hello():
    # 直接读取并返回根目录下的index.html文件
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"读取index.html失败: {e}")
        return "无法加载页面", 500

@app.route("/api/generate", methods=["POST"])
def generate_report():
    """
    这是核心的 API 接口。
    它负责执行完整的"服务串联"。
    """
    print("接收到 /api/generate 请求...")

    try:
        # --- 1. 接收文件和参数 ---
        if 'excel_file' not in request.files:
            return jsonify({"error": "没有找到名为 'excel_file' 的文件"}), 400

        file = request.files['excel_file']
        company_name = request.form.get('company_name', '未知公司')
        report_year = request.form.get('report_year', '2024')

        if file.filename == '':
            return jsonify({"error": "文件名为空"}), 400

        # --- 2. 安全地保存临时文件 ---
        # 我们不能直接用用户上传的文件名，不安全
        filename = secure_filename(file.filename)
        # 我们把它保存在一个临时的、安全的地方
        temp_dir = tempfile.gettempdir()
        temp_excel_path = os.path.join(temp_dir, filename)
        file.save(temp_excel_path)
        print(f"临时文件已保存到: {temp_excel_path}")

        # --- 3. [串联第一步] 调用 DataReader ---
        print("调用 DataReader...")
        reader = ExcelDataReader(temp_excel_path)
        data = reader.extract_data()
        if not data:
            return jsonify({"error": "无法从 Excel 提取数据"}), 500

        # 把 Web 传来的参数也补充进数据字典
        data['company_name'] = company_name
        data['report_year'] = report_year

        # --- 4. [串联第二步] 调用 AIService ---
        print("调用 AIService...")
        # 把从 Excel 读到的数据，交给 AI 去写摘要
        summary_text = ai_service.generate_executive_summary(data)

        # --- 5. [串联第三步] 调用 ReportWriter ---
        print("调用 ReportWriter...")
        writer = WordReportWriter()
        writer.add_title_page(company_name, report_year)

        # 在这里把你新写的摘要加进去
        writer.doc.add_heading("执行摘要", level=1)
        writer.doc.add_paragraph(summary_text)

        # 加入作业一的表格
        writer.add_emission_table(data)

        # --- 6. 准备并返回 Word 文件 ---
        # 我们把 Word 文档也保存在临时目录
        output_filename = "carbon_report_v1.docx"
        temp_word_path = os.path.join(temp_dir, output_filename)
        writer.save(temp_word_path)
        print(f"临时报告已生成: {temp_word_path}")

        # 使用 send_file 把它作为"附件"发回给浏览器
        return send_file(
            temp_word_path,
            as_attachment=True,
            download_name=output_filename # 这是浏览器下载时显示的文件名
        )

    except Exception as e:
        print(f"生成报告时发生严重错误: {e}")
        return jsonify({"error": f"服务器内部错误: {str(e)}"}), 500
    finally:
        # 无论成功还是失败，都尝试清理临时文件
        if 'temp_excel_path' in locals() and os.path.exists(temp_excel_path):
            try:
                os.remove(temp_excel_path)
            except Exception as e:
                print(f"清理临时Excel文件失败: {e}")
        if 'temp_word_path' in locals() and os.path.exists(temp_word_path):
            try:
                os.remove(temp_word_path)
            except Exception as e:
                print(f"清理临时Word文件失败: {e}")
                # 文件可能正在下载，稍后会自动被系统清理