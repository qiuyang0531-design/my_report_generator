# 碳盘查报告生成系统

企业温室气体排放数据自动化报告生成工具，支持从 Excel/CSV 文件提取数据并生成专业的中文 Word 格式报告。

## 项目简介

这是一个碳盘查报告自动生成系统，用于企业温室气体排放核算与报告。系统可以从 Excel 或 CSV 文件中自动提取企业的温室气体排放数据（包括范围一、范围二、范围三排放），通过 AI 生成执行摘要，最终生成符合标准格式的 Word 文档报告。

## 快速开始

### 环境要求

- Python 3.7+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install -r requirements.txt
```

依赖包括：
- `pandas>=2.0.0` - 数据处理
- `openpyxl>=3.1.0` - Excel 文件读取
- `python-docx>=0.8.11` - Word 文档生成
- `docxtpl` - Word 模板渲染
- `openai` - AI 摘要生成（可选）
- `python-dotenv` - 环境变量管理

### 配置环境变量

复制 `.env` 文件并配置（如果需要 AI 摘要功能）：

```bash
OPENAI_API_KEY=your_api_key_here
OPENAI_BASE_URL=your_base_url_here
```

### 运行程序

```bash
python report_generator.py
```

或者指定输入输出路径：

```python
from report_generator import generate_report
generate_report(csv_path="你的数据.csv", output_path="输出报告.docx")
```

### 预期输出

程序运行成功后，会生成 `carbon_report_v1.docx` 文件，包含：
- 封面页（公司名称、报告年度）
- 公司基本信息
- 温室气体排放汇总表
- 范围一、范围二、范围三排放明细
- AI 生成的执行摘要（如配置了 API）

## 文件说明

| 文件 | 作用 |
|------|------|
| `report_generator.py` | **主入口**，一键生成报告的完整流程 |
| `data_reader.py` | 数据提取模块，支持 Excel 和 CSV 文件读取 |
| `report_writer.py` | Word 文档生成模块（备用，当前使用模板方式） |
| `ai_service.py` | AI 摘要生成服务，调用 OpenAI API 生成报告摘要 |
| `template.docx` | Word 报告模板，包含所有变量占位符 |
| `requirements.txt` | Python 依赖包列表 |
| `.env` | 环境变量配置文件 |
| `减排行动统计.csv` | 示例数据文件 |
| `tools/` | 辅助工具目录，包含数据检查和调试脚本 |

### 工具目录

- `find_company_info.py` - 测试数据提取
- `check_emission_data.py` - 检查排放数据完整性
- `analyze_template.py` - 分析模板变量
- `debug_cell_location.py` - 调试 Excel 单元格定位

## 数据源格式

### CSV 文件字段说明

系统支持 CSV 格式作为数据源，主要字段如下：

#### 基本信息
- `company_name` - 企业名称
- `legal_person` - 法定代表人
- `registered_address` - 注册地址
- `date_of_establishment` - 成立日期
- `registered_capital` - 注册资本
- `Unified_Social_Credit_Identifier` - 统一社会信用代码
- `scope_of_business` - 经营范围
- `company_profile` - 公司简介

#### 报告信息
- `reporting_period` - 盘查年度
- `document_number` - 文档编号
- `deadline` - 截止日期
- `evaluation_level` - 评估等级
- `evaluation_score` - 评估分数

#### 排放数据
- `scope_1_emissions` - 范围一排放量（tCO2e）
- `scope_2_location_based_emissions` - 范围二排放（基于位置，tCO2e）
- `scope_2_market_based_emissions` - 范围二排放（基于市场，tCO2e）
- `scope_3_emissions` - 范围三总排放量（tCO2e）
- `scope_3_category_1_emissions` - 范围三分类1排放（外购商品和服务）
- `scope_3_category_2_emissions` - 范围三分类2排放（资本货物）
- `scope_3_category_3_emissions` - 范围三分类3排放（燃料和能源相关）
- `scope_3_category_4_emissions` - 范围三分类4排放（上下游运输）
- `scope_3_category_5_emissions` - 范围三分类5排放（废弃物）
- `scope_3_category_6_emissions` - 范围三分类6排放（员工商务差旅）
- `scope_3_category_7_emissions` - 范围三分类7排放（员工通勤）
- `scope_3_category_9_emissions` - 范围三分类9排放（输入运输配送）
- `scope_3_category_10_emissions` - 范围三分类10排放（产品使用）
- `scope_3_category_12_emissions` - 范围三分类12排放（产品报废处理）

#### 参考信息
- `GWP_Value_Reference_Document` - GWP值参考文档
- `rule_file` - 核算规则文件

### 换企业数据需要修改的地方

1. **替换数据文件**：将新的 CSV 文件替换 `减排行动统计.csv` 或在运行时指定路径

2. **检查字段匹配**：确保新 CSV 包含上述必需字段，字段名需要一致

3. **编码处理**：
   - CSV 文件建议使用 UTF-8 编码
   - 如遇到中文乱码，可在 `data_reader.py:read_emission_data_csv()` 中调整编码参数

4. **自定义模板**（可选）：
   - 修改 `template.docx` 以调整报告格式
   - 模板使用 Jinja2 语法，如 `{{ company_name }}`

5. **区域数据解析**：
   - 如果 CSV 有特定区域划分，需在 `data_reader.py:_parse_csv_sections()` 中适配解析逻辑

## 已知问题

### 1. 动态表格的空行问题

**现象**：生成的 Word 报告中，排放明细表格可能包含多余的空行。

**原因**：
- `data_reader.py` 中 `_parse_csv_sections()` 方法解析 CSV 时，可能读取了空数据行
- 模板渲染时未过滤空值

**修复方法**：
在 `data_reader.py` 约 813-866 行，添加空行过滤：

```python
# 在构建 scope1_items 和 scope2_3_items 后添加过滤
scope1_items = [item for item in scope1_items if item.get('name') and item.get('emission')]
scope2_3_items = [item for item in scope2_3_items if item.get('name') and item.get('emission')]
```

### 2. main.py 引用错误

**现象**：`main.py` 引用了已删除的 `web_api.py`，无法正常运行。

**解决**：使用 `report_generator.py` 作为主入口，或修复 `main.py` 的导入。

### 3. CSV 中文编码问题

**现象**：CSV 文件中的中文字符可能出现乱码。

**解决**：
- 确保 CSV 保存为 UTF-8 编码
- 在 `data_reader.py:read_emission_data_csv()` 中尝试不同编码（`utf-8-sig`, `gbk`, `gb18030`）

### 4. AI 服务依赖

**现象**：未配置 OpenAI API 时，AI 摘要生成会失败。

**解决**：
- 在 `.env` 中配置有效的 API 密钥
- 或修改 `ai_service.py` 中的降级逻辑，在 API 不可用时使用预设模板

### 5. 模板文件路径

**现象**：`template.docx` 文件可能不存在或路径不正确。

**解决**：确保 `template.docx` 在项目根目录，或修改 `report_generator.py` 中的模板路径。

## 开发计划

- [ ] 修复 main.py 入口问题
- [ ] 完善动态表格空行过滤
- [ ] 添加 CSV 编码自动检测
- [ ] 增强 AI 服务降级机制
- [ ] 支持命令行参数指定输入输出
- [ ] 添加单元测试

## 许可证

本项目仅供学习和研究使用。
