import requests
import os

# 确保uploads目录存在
if not os.path.exists('uploads'):
    os.makedirs('uploads')

# 测试文件路径
test_file_path = 'test_data.xlsx'

# 确保测试文件存在
if not os.path.exists(test_file_path):
    print(f"错误: 测试文件 {test_file_path} 不存在")
    exit(1)

# API端点
url = 'http://localhost:5071/api/generate'

# 准备文件和表单数据
files = {
    'excel_file': open(test_file_path, 'rb'),
}
data = {
    'company_name': '测试公司',
    'report_year': '2024'
}

print("开始测试API Key失效时的安全网功能...")
print("正在发送请求生成文档...")

try:
    # 发送请求
    response = requests.post(url, files=files, data=data)
    
    # 检查响应
    if response.status_code == 200:
        # 保存生成的文档
        output_file = 'carbon_report_v1.docx'
        with open(output_file, 'wb') as f:
            f.write(response.content)
        print(f"成功生成文档: {output_file}")
        print("文档已保存，其中的执行摘要应该是安全网生成的模板内容")
    else:
        print(f"请求失败，状态码: {response.status_code}")
        try:
            error_data = response.json()
            print(f"错误信息: {error_data.get('error', '未知错误')}")
        except:
            print(f"响应内容: {response.text}")
            
except Exception as e:
    print(f"发生异常: {str(e)}")
finally:
    # 关闭文件
    files['excel_file'].close()