"""
配置模块 - 数据类和配置定义
=====================

包含所有协议配置的数据结构和常量定义。
"""

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional, Callable


@dataclass
class FieldMapping:
    """字段映射配置"""
    keyword: str           # Excel列关键词
    default: Any          # 默认值
    dtype: str            # 数据类型: 'str', 'float', 'int'
    alt_keywords: List[str] = field(default_factory=list)  # 备选关键词


@dataclass
class TableProtocol:
    """表格协议配置"""
    name: str                     # 协议名称
    description: str              # 描述
    required_keywords: set        # 必需关键词（必须全部匹配）
    optional_keywords: set        # 可选关键词（部分匹配即可）
    field_mappings: Dict[str, FieldMapping]  # 字段映射
    output_var: str               # 输出变量名
    ffill_fields: List[str] = field(default_factory=list)  # 需要前向填充的字段
    min_match_ratio: float = 0.3  # 最小匹配度（可选关键词）
    post_process: Optional[Callable] = None  # 后处理函数
    sheet_name_patterns: List[str] = field(default_factory=list)  # 工作表名称匹配模式
    header_rows_to_check: int = 2  # 表头行检查数量（默认检查2行）


# ==================== 协议优先级配置 ====================
# 优先级高的协议放在前面，先匹配
_PROTOCOL_ORDER = [
    'EmissionFactorProtocol',      # 排放因子表优先级最高（有独特的"低位发热量"+"氧化率"组合）
    'ActivitySummaryProtocol',      # 活动数据汇总表（基于位置）
    'ActivitySummaryMarketProtocol', # 活动数据汇总表（基于市场）
    'GWPProtocol',                  # GWP值表（有独特的"GWP"关键词）
    'GHGInventoryProtocol',         # 温室气体盘查表
    'Scope1EmissionsProtocol',      # 范围一排放源表（最宽泛，放最后）
]
