"""
data_reader 包 - 协议驱动型数据提取器
====================================

核心设计理念：
1. 配置驱动：所有表格识别规则、字段映射、后处理逻辑都由配置定义
2. 解耦识别与解析：通过指纹识别表格类型，通过协议定义提取逻辑
3. 插件化扩展：新增表格类型只需添加配置，无需修改核心代码

架构层次：
- protocols: 协议配置层（定义所有表格类型的识别规则和处理逻辑）
- fingerprint: 表格指纹识别器（负责识别表格类型）
- extractor: 协议提取器（负责根据协议提取数据）
- readers: 专项读取器（负责提取特定类型的数据）
- main: 高层接口（提供统一的数据获取接口）

使用示例（方式1 - 推荐）：
    >>> from data_reader import ExcelDataReaderRefactored
    >>> reader = ExcelDataReaderRefactored("test_data.xlsx")
    >>> context = reader.get_all_context()
    >>> reader.close()

子模块导入（高级用法）：
    >>> from data_reader.protocols import TABLE_PROTOCOLS
    >>> from data_reader.readers import Scope1Reader
"""

# 主要导出
from .main import ExcelDataReaderRefactored

# 子模块导出（供高级用法）
from . import config
from . import protocols
from . import utils
from . import post_processors
from . import fingerprint
from . import extractor
from . import readers

__version__ = '2.0.0'

__all__ = [
    'ExcelDataReaderRefactored',
    # 子模块（可选导出）
    'config',
    'protocols',
    'utils',
    'post_processors',
    'fingerprint',
    'extractor',
    'readers',
]
