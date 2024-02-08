""" 
author: kefengpl
email: xu2332700545@gmail.com
last update: Mon Jan 29 18:45:55 2024 +0800

文件是实现用户交互的接口，用以设置数据文件路径，获得基金产品报告。
设计原则：简化交互使用，
"""
from report_generate import *

def multi_fund_report_interface():
    """ 在基金对的标指数数据一致的情况下，允许同时导出多个基金的基金报告。"""
    netval_path = "data/中证1000基金-净值数据.xlsx" # 净值数据表路径，可以包含多个基金，每个基金的数据是一列，列名是基金名称。不要出现任何冗余列。
    index_path = "data/裕锦中证1000指数增强-指数数据.xlsx"  # 所有基金共用该指数数据表。指数数据表路径：① 非指增：允许添加多列指数数据；② 指增：只允许添加一列指数数据
    enhanced_fund: bool = False # 指定是否为指增基金：注意：指增基金和非指增基金不能同时批量计算
    start_dates: list = [] # 从净值数据从前往后按顺序指定起始计算日期。如果有三个基金，仅指定前两个日期，可以写为 [date(2022, 8, 19), date(2022, 9, 19)]
                     # 如果有三个基金，需要指定第一个基金和第三个基金的开始计算日期，可以写为 [date(2022, 8, 19), None, date(2022, 9, 19)]
    corp_names: list = ["裕锦", "图灵"] # 从前往后按顺序匹配公司名称，使用方法与 start_dates 的使用完全一致
    add_indicators_tables: bool = True # 添加 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表
    multi_fund_report(netval_path, index_path, enhanced_fund, corp_names = corp_names, 
                      start_dates = start_dates, add_indicators_tables = add_indicators_tables)

def main():
    multi_fund_report_interface()

if __name__ == "__main__":
    main()