# 协议驱动型数据提取器 - 重构指南

## 概述

重构后的 `data_reader_refactored.py` 采用了**协议驱动型架构**，实现了以下目标：

1. **配置驱动**：所有表格识别规则、字段映射、后处理逻辑都由配置定义
2. **解耦识别与解析**：通过指纹识别表格类型，通过协议定义提取逻辑
3. **插件化扩展**：新增表格类型只需添加配置，无需修改核心代码

## 架构层次

```
TABLE_PROTOCOLS (协议配置层)
    ↓
TableFingerprint (表格指纹识别器)
    ↓
ProtocolExtractor (协议数据提取器)
    ↓
ExcelDataReaderRefactored (高层数据读取器)
```

## 核心组件

### 1. TableProtocol (协议配置)

```python
TableProtocol(
    name='排放因子表',                    # 协议名称
    description='...',                   # 描述
    required_keywords={'低位发热量', '氧化率'},  # 必需关键词
    optional_keywords={...},              # 可选关键词
    field_mappings={...},                 # 字段映射
    output_var='pro_ef_items',            # 输出变量名
    ffill_fields=['category'],            # 需要前向填充的字段
    min_match_ratio=0.3,                  # 最小匹配度
    sheet_name_patterns=['附表2-EF'],     # 工作表名称模式
)
```

### 2. TableFingerprint (指纹识别器)

负责识别工作表类型，支持两种识别方式：
- **关键词匹配**：通过表头内容识别
- **名称匹配**：通过工作表名称识别

### 3. ProtocolExtractor (数据提取器)

负责根据协议配置提取数据：
- 查找表头行
- 构建列映射
- 提取数据行
- 应用前向填充
- 执行后处理

### 4. ExcelDataReaderRefactored (高层数据读取器)

提供统一的接口获取所有数据：
```python
reader = ExcelDataReaderRefactored("test_data.xlsx")
context = reader.get_all_context()  # 一键获取所有数据
reader.close()
```

## 使用示例

### 基本使用

```python
from data_reader_refactored import ExcelDataReaderRefactored

# 创建读取器
reader = ExcelDataReaderRefactored("test_data.xlsx")

# 获取所有数据
context = reader.get_all_context()

# 访问提取的数据
pro_ef_items = context.get('pro_ef_items', [])
gwp_items = context.get('gwp_items', [])

# 关闭读取器
reader.close()
```

### 获取特定协议的数据

```python
# 获取排放因子表数据
ef_data = reader.get_protocol_data('EmissionFactorProtocol')

# 获取GWP值表数据
gwp_data = reader.get_protocol_data('GWPProtocol')
```

## 添加新表格类型

要添加新的表格类型，只需在 `TABLE_PROTOCOLS` 中添加配置：

```python
TABLE_PROTOCOLS['NewTableProtocol'] = TableProtocol(
    name='新表格类型',
    description='表格描述',
    required_keywords={'关键词1', '关键词2'},
    optional_keywords={'可选关键词1', '可选关键词2'},
    field_mappings={
        'field1': FieldMapping('列名1', default_value, 'str'),
        'field2': FieldMapping('列名2', 0, 'float'),
    },
    output_var='new_table_items',
    ffill_fields=[],
    min_match_ratio=0.3,
)
```

## 协议优先级

协议按照 `_PROTOCOL_ORDER` 中定义的顺序进行匹配：

```python
_PROTOCOL_ORDER = [
    'EmissionFactorProtocol',    # 排放因子表（优先级最高）
    'ActivitySummaryProtocol',   # 活动数据汇总表
    'GWPProtocol',                # GWP值表
    'GHGInventoryProtocol',       # 温室气体盘查表
    'Scope1EmissionsProtocol',    # 范围一排放源表（最宽泛）
]
```

## 与原代码的对比

| 特性 | 原代码 | 重构后代码 |
|------|--------|-----------|
| 代码行数 | ~900行 | ~550行 |
| 协议数量 | 分散在多个方法中 | 集中在TABLE_PROTOCOLS配置 |
| 添加新表格 | 需要修改多个方法 | 只需添加配置 |
| 识别逻辑 | 硬编码 | 配置驱动 |
| 可维护性 | 低 | 高 |
| 可扩展性 | 低 | 高 |

## 迁移指南

### 1. 替换导入

```python
# 旧代码
from data_reader import ExcelDataReader

# 新代码
from data_reader_refactored import ExcelDataReaderRefactored as ExcelDataReader
```

### 2. API兼容性

新代码保持与原代码相同的API接口：

```python
reader = ExcelDataReader("test_data.xlsx")
context = reader.get_all_context()  # 或者使用 reader.extract_data_from_xlsx_dynamic()
```

### 3. 输出变量

新代码使用相同的输出变量名，确保与模板兼容：
- `pro_ef_items` / `emission_factor_items`
- `gwp_items`
- `ghg_inventory_items`
- `activity_summary_items`
- `scope1_*_emissions_items`

## 技术亮点

1. **数据类（dataclass）**：使用Python 3.7+的数据类简化配置
2. **类型提示**：完整的类型注解提高代码可读性
3. **优先级匹配**：避免歧义，确保最准确的识别
4. **工作表名称匹配**：支持基于名称的快速识别
5. **自动前向填充**：配置驱动的合并单元格处理
6. **后处理钩子**：支持自定义后处理逻辑
