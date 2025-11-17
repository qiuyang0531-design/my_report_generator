from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
import os
from dotenv import load_dotenv
import tempfile
from datetime import datetime
import uuid

# 导入现有的模块
from data_reader import ExcelDataReader
from report_writer import WordReportWriter
from ai_service import AIService

# 加载环境变量
load_dotenv()

# 创建Flask应用实例
app = Flask(__name__)

# 初始化AI服务
ai_service = AIService()

# 配置文件上传设置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """检查文件是否是允许的类型"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """首页路由"""
    return render_template('index.html')

@app.route('/generate_report', methods=['POST'])
def generate_report():
    """生成报告的路由"""
    # 检查是否有文件被上传
    if 'file' not in request.files:
        flash('请选择一个文件上传')
        return redirect(request.url)
    
    file = request.files['file']
    
    # 如果用户没有选择文件，浏览器也会提交一个空的文件部分
    if file.filename == '':
        flash('请选择一个文件')
        return redirect(request.url)
    
    # 确保文件类型正确
    if file and allowed_file(file.filename):
        # 生成唯一的文件名以避免冲突
        unique_id = str(uuid.uuid4())[:8]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{unique_id}_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # 使用ExcelDataReader读取数据
            reader = ExcelDataReader(filepath)
            report_data = reader.extract_all_data()
            
            # 使用AI服务生成执行摘要
            try:
                executive_summary = ai_service.generate_executive_summary(report_data)
                report_data['executive_summary'] = executive_summary
                print("执行摘要已添加到报告数据中")
            except Exception as e:
                print(f"添加执行摘要时出错: {e}")
                # 即使AI摘要失败，程序也继续运行
            
            # 使用WordReportWriter生成报告
            writer = WordReportWriter(
                template_path=os.getenv('TEMPLATE_PATH', '模板1.docx'),
                cover_image_path=os.getenv('COVER_IMAGE_PATH', '封面.png')
            )
            
            # 生成报告文件路径
            report_filename = "carbon_report_v1.docx"
            report_filepath = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
            
            # 写入报告数据
            writer.write_report(report_data, report_filepath)
            
            # 保存报告路径到session，以便下载时使用
            app.config['LAST_REPORT_PATH'] = report_filepath
            app.config['LAST_REPORT_FILENAME'] = report_filename
            
            # 返回结果页面，包含下载链接
            return render_template('result.html', data=report_data, show_download=True)
            
        except Exception as e:
            error_msg = f"处理文件时出错: {str(e)}"
            print(error_msg)  # 在控制台记录详细错误
            flash(error_msg)
            return redirect('/')
        finally:
            # 清理上传的Excel文件，但保留生成的报告文件
            if os.path.exists(filepath):
                os.remove(filepath)
    
    flash('不支持的文件类型，请上传.xlsx格式的Excel文件')
    return redirect('/')

@app.route('/download_report')
def download_report():
    """下载生成的报告文件"""
    report_path = app.config.get('LAST_REPORT_PATH')
    report_filename = app.config.get('LAST_REPORT_FILENAME')
    
    if report_path and os.path.exists(report_path):
        try:
            # 发送文件给用户下载
            return send_file(
                report_path,
                as_attachment=True,
                download_name=report_filename,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        except Exception as e:
            print(f"下载文件时出错: {str(e)}")
            flash('下载报告时出错')
            return redirect('/')
    else:
        flash('没有找到可下载的报告文件')
        return redirect('/')
    
    return "不支持的文件类型"

# 设置Flask密钥用于session和flash消息
app.secret_key = os.getenv('SECRET_KEY', 'dev_key_change_in_production')

# 设置最大上传文件大小
app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_CONTENT_LENGTH', 16777216))  # 默认16MB

if __name__ == '__main__':
    # 从环境变量获取配置
    host = os.getenv('HOST', '0.0.0.0')
    port = int(os.getenv('PORT', 5071))
    debug = os.getenv('FLASK_ENV') == 'development' or os.getenv('DEBUG', 'False').lower() == 'true'
    
    # 运行服务器
    print(f"服务器启动在 http://{host}:{port}")
    print(f"调试模式: {debug}")
    app.run(debug=debug, host=host, port=port)