"""
报告配置模块 - 专业碳核算版
包含 22 个排放源的精细化量化方法说明
"""

class ReportConfig:
    def __init__(self, company_name="某公司", reporting_period="2024年度"):
        self.company_name = company_name
        self.reporting_period = reporting_period

    def get_quantification_methods(self):
        c = self.company_name
        p = self.reporting_period

        return {
            'scope_1': {
                # (一) 固定燃烧：天然气
                'natural_gas': {
                    'name': '热处理炉/加热炉（天然气燃烧）',
                    'model': '来源于《2006 年 IPCC 国家温室气体清单指南》第2卷第2章公式2.1',
                    'ad': f'来源于{c}提供的《天然气发票》，统计周期为{p}。',
                    'ef': '参数包括天然气低位发热量、单位热值含碳量、碳氧化率，来源于《钢铁生产核算指南》附表A.1。'
                },
                # (二) 固定燃烧：液化石油气
                'lpg': {
                    'name': '冶钢宾馆（液化石油气燃烧）',
                    'model': '来源于《2006 年 IPCC 国家温室气体清单指南》第2卷第2章公式2.1',
                    'ad': f'来源于{c}提供的《{p}冶钢宾馆水电气》每月使用液化石油气的抄表数据统计。',
                    'ef': '参数来源于《企业温室气体排放核算与报告指南-钢铁生产》附表A.1常用化石燃料相关参数缺省值。'
                },
                # (三) 固定燃烧：固体燃料
                'solid_fuels': {
                    'name': '固体燃料（无烟煤、烟煤、焦炭等燃烧）',
                    'model': '方法学来自标准ISO 14064-1:2018/ 6.2',
                    'ad': f'来源于{c}提供《物资收发存汇总表》在{p}内的发出量（干重），并根据含水率求得收到基重量。',
                    'ef': '参数包括低位发热量、单位热值含碳量等，来源于《钢铁生产核算指南》附表A.1缺省值。'
                },
                # (四) 副产品外售
                'byproducts': {
                    'name': '副产品外售（焦炉煤气、粗苯）',
                    'model': '方法学来自标准ISO 14064-1:2018/ 6.2',
                    'ad': f'来源于{c}提供《2024年能源报表》及《新化能指标完成情况》，统计周期为{p}。',
                    'ef': '参数来源于《企业温室气体排放核算与报告指南-钢铁生产》附表A.1。'
                },
                # (五) 移动燃烧：柴油
                'diesel_transport': {
                    'name': '场内运输（0#柴油燃烧）',
                    'model': '来源于《2006 年 IPCC 国家温室气体清单指南》第2卷第2章公式2.1',
                    'ad': f'来源于{c}提供《物资收发存汇总表》在{p}内场内运输使用0#柴油的统计。',
                    'ef': '来源于《GB/T 2589-2020》及《IPCC-2006缺省值》，氧化率保守取值100%。'
                },
                # (六) 移动燃烧：汽油
                'gasoline_transport': {
                    'name': '公务车（92#汽油燃烧）',
                    'model': '来源于《2006 年 IPCC 国家温室气体清单指南》第2卷第2章公式2.1',
                    'ad': f'来源于{c}提供《总经办小车队燃油表》，统计周期为{p}。',
                    'ef': '来源于《GB/T 2589-2020》及《IPCC-2006缺省值》，氧化率保守取值100%。'
                },
                # (七) 散逸排放：空调制冷剂
                'ac_refrigerant': {
                    'name': '分体式空调/汽车空调（R32）',
                    'model': '来源于IPCC《2006 年国家温室气体清单指南》第3卷第7章公式7.13',
                    'ad': f'来源于{c}各部门提供的《空调加氟汇总表》，统计周期为{p}。',
                    'ef': '来源于2006年IPCC国家温室气体清单指南V2_3_Ch3 Table3.2.1/3.2.2。'
                },
                # (八) 散逸排放：灭火器
                'fire_extinguisher': {
                    'name': 'CO2灭火器灭火系统逸散',
                    'model': '来源于IPCC《2006 年国家温室气体清单指南》V3_7_Ch7 7.6.2.2',
                    'ad': f'来源于{c}提供的《物资收发存汇总表》中在{p}内的灭火剂采购量数据。',
                    'ef': '气体灭火系统因子取2%，便携式CO2灭火器因子取100%（IPCC 2006缺省值）。'
                },
                # (九) 逸散排放：化粪池
                'septic_tank': {
                    'name': '化粪池（CH4逸散）',
                    'model': '来源于IPCC《2006 年国家温室气体清单指南》第5卷第6章公式6.1、6.3',
                    'ad': f'根据{c}在{p}内的员工出勤总工时推算BOD排放量。',
                    'ef': 'EF=Bo*MCF=0.06（kgCH4/kgBOD），数据源自IPCC 2006第5卷表6.2/6.3/6.4。'
                },
                # (十) 逸散排放：厌氧池
                'anaerobic_pond': {
                    'name': '厌氧池（CH4逸散）',
                    'model': '来源于IPCC《2006 年国家温室气体清单指南》第5卷第6章公式6.1、6.4',
                    'ad': f'来源于{c}提供《细化收资清单（动力）（六氟化硫）》，统计周期为{p}。',
                    'ef': 'EF=Bo*MCF=0.025（kgCH4/kgCOD），Bo取0.25，MCF取0.1。'
                },
                # (十一) 逸散排放：SF6开关
                'sf6_switch': {
                    'name': '高压开关（SF6逸散）',
                    'model': '该方法学来自标准ISO14064-1：2018/ 6.2',
                    'ad': f'来源于{c}提供《细化收资清单（动力）（六氟化硫）》高压开关填充SF6铭牌额定量。',
                    'ef': '根据GB/T 8905-2012《六氟化硫电气设备中气体管理和检测导则》逸散系数取0.5%。'
                },
                # (十二) 工业过程排放：制程排放
                'process_co2': {
                    'name': '制程排放（CO2排放）',
                    'model': '来源于《2006 年 IPCC 国家温室气体清单指南》第2卷第2章公式2.1',
                    'ad': f'来源于{c}提供《2024年统计月报》中石灰石、白云石、增碳剂、合金等原料在{p}内的发出量。',
                    'ef': '来源于《企业温室气体排放核算与报告指南-钢铁生产》附表A.2及A.3缺省值。'
                }
            },
            'scope_2': {
                # (十三) 外购电力
                'electricity': {
                    'name': '外购电力（CO2排放）',
                    'model': '来源于《工业其他行业企业温室气体排放核算方法与报告指南》公式14',
                    'ad': f'来源于{c}提供的《电费发票》及《电力消费凭证》，统计周期为{p}。',
                    'ef': '来源于湖北省2022年省级电力排放因子及全国电力平均排放因子。'
                }
            },
            'scope_3': {
                # (十四) 类别1
                'category_1': {
                    'name': '外购商品和服务',
                    'model': '来源于《Technical Guidance for Calculating Scope 3 Emissions》的Average-data method和Spend-based method',
                    'ad': f'来源于{c}提供《物资收发存汇总表》入库量，涵盖铁矿石、生石灰、废钢等原材料。',
                    'ef': '来源于ecoinvent 3.10-cut off及SupplyChainGHGEmissionFactors数据库。'
                },
                # (十五) 类别2
                'category_2': {
                    'name': '资本货物',
                    'model': '来源于《Scope 3计算指南》的Spend-based method',
                    'ad': f'来源于{c}提供《物资收发存汇总表》入库量中空调、电子计算机、工业设备等采购花费金额。',
                    'ef': '来源于SupplyChainGHGEmissionFactors_v1.3.0数据库。'
                },
                # (十六) 类别3
                'category_3': {
                    'name': '燃料和能源相关的活动',
                    'model': '来源于《Scope 3计算指南》Category 3的Average-data method',
                    'ad': f'来源于{c}提供的能源报表，涵盖液化气、煤炭、柴油、汽油、电力在{p}内的生产侧排放。',
                    'ef': '来源于ecoinvent 3.10-cut off及2023年全国电力平均碳足迹因子。'
                },
                # (十七) 类别4
                'category_4': {
                    'name': '上游运输和配送',
                    'model': '来源于《Scope 3计算指南》Category 4的Distance-based method',
                    'ad': f'根据{c}提供供应商运输距离及产品销售统计，加权计算{p}内的运输排放。',
                    'ef': '船运、汽运、铁路运输因子来源于ecoinvent 3.10-cut off。'
                },
                # (十八) 类别5
                'category_5': {
                    'name': '运营中产生的废弃物',
                    'model': '来源于《Scope 3计算指南》Category 5的Waste-type-specific method',
                    'ad': f'来源于{c}提供《垃圾表》及《环保指标统计》中{p}内产生的生活垃圾、废油等总量。',
                    'ef': '处理及运输因子来源于ecoinvent 3.10-cut off。'
                },
                # (十九) 类别6
                'category_6': {
                    'name': '员工商务差旅',
                    'model': '来源于《Scope 3计算指南》Category 6',
                    'ad': f'来源于{c}提供《差旅情况统计》中{p}内所有的航运、铁路、住宿人次数据。',
                    'ef': '因子来源于CPCD数据库及ecoinvent 3.10-cut off。'
                },
                # (二十) 类别7
                'category_7': {
                    'name': '员工上下班通勤',
                    'model': '来源于《Scope 3计算指南》Category 7',
                    'ad': f'来源于{c}提供《员工通勤统计》，按平均通勤距离及交通方式进行估算。',
                    'ef': '因子来源于CPCD数据库及ecoinvent 3.10-cut off。'
                },
                # (二十一) 类别10
                'category_10': {
                    'name': '已售产品的加工',
                    'model': '来源于《Scope 3计算指南》Category 10: Processing of sold Products',
                    'ad': f'来源于{c}提供{p}内《钢坯钢材外销情况统计表》中的外销总量。',
                    'ef': '选用三环锻造、郑煤机智鼎等代表性加工强度加权平均所得。'
                },
                # (二十二) 类别12
                'category_12': {
                    'name': '已售产品的报废处理',
                    'model': '来源于《Scope 3计算指南》Category 12: End-of-Life Treatment',
                    'ad': f'根据{c}提供的外销总量，结合全国废钢报废比例进行推算。',
                    'ef': '因子来源于ecoinvent 3.10-cut off。'
                },
                # (十六) 类别8
                'category_8': {
                    'name': '上游租赁资产',
                    'model': '来源于《Scope 3计算指南》Category 8',
                    'ad': f'{c}的上游租赁资产主要包括租赁的办公场所、设备等，该类别数据不具备重要性，不进行量化。',
                    'ef': '租赁资产能源消耗因子来源于ecoinvent 3.10-cut off。'
                },
                # (十七) 类别9
                'category_9': {
                    'name': '下游运输和配送',
                    'model': '来源于《Scope 3计算指南》Category 9: Downstream Transportation',
                    'ad': f'来源于{c}提供《销售数据》中{p}内产品从公司运输到客户产生的排放。',
                    'ef': '运输因子来源于ecoinvent 3.10-cut off。'
                },
                # (十八) 类别11
                'category_11': {
                    'name': '已售产品的使用',
                    'model': '来源于《Scope 3计算指南》Category 11: Use of Sold Products',
                    'ad': f'来源于{c}提供的产品使用阶段能源消耗数据，涵盖{p}内已售产品的使用过程排放。',
                    'ef': '使用阶段排放因子来源于ecoinvent 3.10-cut off。'
                },
                # (十九) 类别13
                'category_13': {
                    'name': '下游租赁资产',
                    'model': '来源于《Scope 3计算指南》Category 13: Downstream Leased Assets',
                    'ad': f'{c}的下游租赁资产主要包括租给客户的设备、场所等，该类别数据不具备重要性，不进行量化。',
                    'ef': '租赁资产能源消耗因子来源于ecoinvent 3.10-cut off。'
                },
                # (二十) 类别14
                'category_14': {
                    'name': '特许经营',
                    'model': '来源于《Scope 3计算指南》Category 14: Franchises',
                    'ad': f'{c}的特许经营排放主要包括特许经营店运营产生的排放，该类别数据不具备重要性，不进行量化。',
                    'ef': '特许经营店排放因子来源于ecoinvent 3.10-cut off。'
                },
                # (二十一) 类别15
                'category_15': {
                    'name': '投资',
                    'model': '来源于《Scope 3计算指南》Category 15: Investments',
                    'ad': f'{c}的投资排放主要包括对被投资公司的股权投资按比例分摊的排放，该类别数据不具备重要性，不进行量化。',
                    'ef': '投资排放因子来源于ecoinvent 3.10-cut off。'
                }
            }
        }

    def get_scope_3_category_name(self, category_num):
        """
        获取范围三全部 15 个类别名称

        Args:
            category_num: 类别编号 (1-15)

        Returns:
            类别名称字符串
        """
        names = {
            1: "外购商品和服务",
            2: "资本货物",
            3: "燃料和能源相关活动",
            4: "上游运输和配送",
            5: "运营中产生的废弃物",
            6: "员工商务差旅",
            7: "员工上下班通勤",
            8: "上游租赁资产",
            9: "下游运输和配送",
            10: "已售产品的加工",
            11: "已售产品的使用",
            12: "已售产品的报废处理",
            13: "下游租赁资产",
            14: "特许经营",
            15: "投资"
        }
        return names.get(category_num, f"类别{category_num}")

    def get_all_scope_3_category_names(self):
        """
        获取范围三全部 15 个类别名称的字典

        Returns:
            字典，键为类别编号，键值为类别名称
        """
        return {
            f'category_{i}': self.get_scope_3_category_name(i)
            for i in range(1, 16)
        }
