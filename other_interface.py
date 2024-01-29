""" 此文件是其它接口，为了简化交互的使用难度，一些可能用不到的接口就写到这里 """

from report_generate import *

def single_fund_report_interface():
    """ 获得基金分析报告 """
    netval_path = "data/裕锦中证1000指数增强-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    index_path = "data/裕锦中证1000指数增强-指数数据.xlsx"  # 指数数据表路径：① 非指增：允许添加多列指数数据；② 指增：只允许添加一列指数数据
    enhanced_fund = True # 需要在参数中手动指明是否是指增类基金，False表示不是指增基金，True表示指增基金
    corp_name = "裕锦量化" # 私募基金管理人名称，不想写可以删了
    start_date = date(2022, 8, 19)  # 指定开始计算的日期，当然，该参数可以删除，这行也可以删除或者赋值为 None，默认从净值数据的起始日期开始计算。
    add_indicators_tables: bool = True # 添加 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表
    single_fund_report(netval_path, index_path, enhanced_fund, corp_name, start_date = start_date,
                       add_indicators_tables = add_indicators_tables)
    
def single_fund_indicator_tables_interface():
    """ 生成单个基金产品各类指标(不包括月度/年度指标，因为它们已经汇总在 fund_report中了)汇总表，滚动收益统计表，放在word中 """
    netval_path = "data/图灵进取中证1000指数增强-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    corp_name = "图灵量化" # 私募基金管理人名称，不想写可以删了
    single_fund_indicator_tables(netval_path, corp_name)  

def multi_fund_indicator_tables_interface():
    """ 生成多个基金产品各类指标(不包括月度/年度指标，因为它们已经汇总在 fund_report中了)汇总表，滚动收益统计表，放在word中 """
    netval_path = "data/中证1000基金-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    corp_names: list = ["裕锦私募", "图灵量化"] # 私募基金管理人名称列表，不想写可以删了
    start_dates: list = [] # 可以指定每只基金的起始计算日期，必须是date格式，例如 date(2022, 8, 19)
    multi_fund_indicator_tables(netval_path, corp_names = corp_names, start_dates = start_dates)  