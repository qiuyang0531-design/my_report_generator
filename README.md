# 碳盘查报告生成系统

这是一个自动化的碳排放报告生成工具，能够从Excel数据中提取温室气体排放信息，并生成格式化的Word报告文档。

## 项目简介

本系统专为中文企业环境设计，支持从标准格式的Excel碳盘查数据表中自动提取以下信息：

- 公司名称
- 报告年份
- 范围一排放（Scope 1）
- 范围二排放（Scope 2）
- 范围三排放（Scope 3）
- 总排放量（基于位置/基于市场）

然后自动生成包含封面页、排放汇总表等专业格式的Word报告文档。

## 系统要求

- Python 3.8 或更高版本
- pip 包管理工具

## 安装步骤

1. 克隆或下载项目到本地

2. 安装依赖包：

```bash
pip install -r requirements.txt
```

3. 验证安装：

```bash
python main.py
```

## 如何使用

### 基本用法

1. **准备数据文件**：

   将你的Excel数据文件命名为 `test_data.xlsx` 放入项目根目录，或在 `main.py` 中修改 `INPUT_EXCEL_PATH` 常量指向你的文件。

   Excel文件需要包含：
   - 公司名称（在文件中搜索"组织名称"）
   - 报告年份（在文件中搜索"盘查年度"）
   - 温室气体排放数据表

2. **运行程序**：

```bash
python main.py
```

3. **查看报告**：

   程序运行完成后，会在项目根目录生成 `碳盘查报告.docx` 文件。

### 文件说明

- **输入文件**：`test_data.xlsx` - 包含碳排放数据的Excel文件
- **输出文件**：`碳盘查报告.docx` - 生成的Word报告
- **封面图片**：`封面.png` - 报告封面使用的图片
- **模板文件**：`模板1.docx` - 可选的Word模板参考

### 自定义配置

编辑 `main.py` 文件中的常量：

```python
INPUT_EXCEL_PATH = 'test_data.xlsx'      # 输入Excel文件路径
OUTPUT_REPORT_PATH = '碳盘查报告.docx'   # 输出Word文件路径
```

## 项目结构

```
my_report_generator/
├── main.py                 # 主程序入口
├── data_reader.py          # Excel数据读取模块
├── report_writer.py        # Word报告生成模块
├── requirements.txt        # Python依赖包列表
├── test_data.xlsx         # 示例输入数据
├── 封面.png               # 报告封面图片
├── 模板1.docx             # 参考Word模板
├── README.md              # 项目说明文档
├── CLAUDE.md              # Claude Code开发指南
├── .gitignore             # Git忽略文件配置
└── tools/                 # 辅助工具脚本
    ├── find_company_info.py
    ├── check_emission_data.py
    ├── analyze_template.py
    ├── debug_cell_location.py
    └── check_sheets.py
```

## 核心功能

### 1. Excel数据读取（data_reader.py）

- **智能关键词搜索**：自动在Excel中查找公司信息和排放数据
- **多方法数据定位**：支持右侧、下方、对角线等多种搜索策略
- **数据验证**：验证提取的数据完整性和准确性
- **错误处理**：详细的错误报告和调试信息

主要数据提取：
- 公司名称（搜索关键词："组织名称"）
- 报告年份（搜索关键词："盘查年度"）
- 范围一排放（Scope 1）
- 范围二排放（Scope 2，支持位置基准和市场基准）
- 范围三排放（Scope 3）
- 总排放量（位置基准和市场基准）

### 2. Word报告生成（report_writer.py）

- **专业格式**：遵循中文商务报告标准格式
- **样式化排版**：使用Word样式实现内容与样式分离
- **中文支持**：内置中文字体支持（宋体、黑体）
- **封面生成**：自动添加封面页和公司信息
- **数据表格**：生成格式化排放汇总表

报告包含：
- 封面页（含封面图片）
- 公司信息和报告年份
- 排放数据汇总表（范围1、2、3和总计）

### 3. 主工作流（main.py）

- **完整流程**：从数据读取到报告生成的一站式处理
- **错误处理**：全面的异常捕获和错误报告
- **用户反馈**：详细的进度提示和处理结果
- **日志记录**：生成详细日志供调试使用

## 辅助工具

### 数据分析工具

`tools/` 目录下提供多个实用脚本：

1. **find_company_info.py** - 测试公司信息提取
2. **check_emission_data.py** - 验证排放数据定位
3. **analyze_template.py** - 分析Word模板结构
4. **debug_cell_location.py** - 调试Excel单元格位置
5. **check_sheets.py** - 检查Excel工作表信息

使用方法：

```bash
python tools/find_company_info.py
python tools/check_emission_data.py
python tools/analyze_template.py
```

## 技术栈

### 核心依赖

- **pandas >= 2.0.0** - 数据处理和分析
- **openpyxl >= 3.1.0** - Excel文件读写支持
- **python-docx >= 0.8.11** - Word文档生成

### 开发工具

- Python 3.9+
- Git版本控制
- Visual Studio Code（推荐）

## 开发指南

### 代码结构

项目采用面向对象设计，核心模块：

- **ExcelDataReader** 类（data_reader.py）：负责Excel数据读取和提取
- **WordReportWriter** 类（report_writer.py）：负责Word报告生成
- **main_workflow()** 函数（main.py）：协调完整工作流程

### 扩展功能

如需扩展报告内容，可以：

1. 在 `report_writer.py` 中添加新的章节方法
2. 在 `main.py` 中调用新添加的方法
3. 参考现有章节实现，如 `add_title_page()` 和 `add_emission_table()`

示例：

```python
# 在 report_writer.py 中添加
class WordReportWriter:
    def add_analysis_section(self, data):
        """添加分析章节"""
        # 实现分析章节内容
        pass

# 在 main.py 中调用
writer.add_analysis_section(data)
```

## 故障排除

### 常见问题

1. **找不到输入文件**
   - 确保 `test_data.xlsx` 存在于项目根目录
   - 检查 `main.py` 中的 `INPUT_EXCEL_PATH` 配置

2. **数据提取失败**
   - 使用 `tools/check_emission_data.py` 检查数据结构
   - 确认Excel文件包含必要的标识（如"组织名称"、"盘查年度"）
   - 检查 `data_reader.py` 中的关键词是否匹配

3. **Word生成错误**
   - 确认 `封面.png` 图片文件存在
   - 检查是否有写入权限

### 技术文档

更详细的技术文档请参见 `CLAUDE.md`。

## 许可证

本项目仅供学习和内部使用。

## 贡献

欢迎提交问题和改进建议！

## 版本历史

- v1.0.0 - 初始版本，支持基本报告生成

## 联系方式

如有问题或建议，请通过项目仓库联系。
