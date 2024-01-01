""" 文件是实现用户交互的接口，用以设置数据文件路径，获得基金产品报告 """

import pandas as pd
from report_generate import *

def single_fund_report(netval_path: str, index_path: str, enhanced_fund: bool, corp_name: str = "私募管理人"):
    """
    生成一只基金的报告，当然，在本函数体内写循环即可将它改为同时生成多只基金的报告。输出的文件在 output 文件夹下

    Args:
        - netval_path (str): 净值数据路径，数据表第一列必须是日期，第二列必须是净值数据，且净值数据列名必须等于产品名
        - index_path (str): 指数数据文件路径，数据表第一列必须是日期，后面的列可以传入多个指数数据(允许传入收盘价)，每列列名必须等于指数名
        - corp_name (str): 私募管理人名称，如果没有输入该参数，默认是 "私募管理人"
        - enhanced_fund (bool): 是否是指增基金
    """
    netval_data = pd.read_excel(netval_path, index_col = 0)
    index_data = pd.read_excel(index_path, index_col = 0)
    fund_name = netval_data.columns[0]
    index_name = index_data.columns[0]
    generate_report(netval_data.iloc[:, 0], index_data, enhanced_fund, corp_name,
                    fund_name = fund_name, index_name = index_name)

def main():
    netval_path = "data/图灵进取中证1000指数增强-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    index_path = "data/图灵进取中证1000指数增强-指数数据.xlsx"  # 指数数据表路径：① 非指增：允许添加多列指数数据；② 指增：只允许添加一列指数数据
    enhanced_fund = True # 需要在参数中手动指明是否是指增类基金，False表示不是指增基金，True表示指增基金
    corp_name = "图灵基金" # 私募基金管理人名称
    single_fund_report(netval_path, index_path, enhanced_fund, corp_name)

if __name__ == "__main__":
    main()