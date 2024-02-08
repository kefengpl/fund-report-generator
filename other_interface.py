""" 此文件是其它接口，为了简化交互的使用难度，一些可能用不到的接口就写到这里。(可以将该文件直接废弃，删除该文件不影响正常运行) """

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

def single_fund_report(netval_path: str, index_path: str, enhanced_fund: bool, corp_name: str = "私募管理人", **kwargs):
    """
    生成一只基金的报告，当然，在本函数体内写循环即可将它改为同时生成多只基金的报告。输出的文件在 output 文件夹下

    Args:
        - netval_path (str): 净值数据路径，数据表第一列必须是日期，第二列必须是净值数据，且净值数据列名必须等于产品名
        - index_path (str): 指数数据文件路径，数据表第一列必须是日期，后面的列可以传入多个指数数据(允许传入收盘价)，每列列名必须等于指数名
        - corp_name (str): 私募管理人名称，如果没有输入该参数，默认是 "私募管理人"
        - enhanced_fund (bool): 是否是指增基金
        - start_date (date, optional): 可选参数，起始计算日期，可以不填，如果填写必须填 datetime.date 格式. 
        - add_indicators_tables (bool, optional): 可选参数，表示是否包含 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表

    """
    netval_data = pd.read_excel(netval_path, index_col = 0)
    index_data = pd.read_excel(index_path, index_col = 0)
    fund_name = netval_data.columns[0]
    index_name = index_data.columns[0] # 注意：index_name 仅适用于指增基金
    start_date: date = kwargs.get("start_date", None)
    add_indicators_tables: bool = kwargs.get("add_indicators_tables", False)
    generate_report(netval_data.iloc[:, 0], index_data, enhanced_fund, corp_name, start_date,
                    add_indicators_tables = add_indicators_tables, fund_name = fund_name, index_name = index_name)
    
def single_fund_indicator_tables(netval_path: str, corp_name: str = utils.CORP_DEFAULT_NAME, **kwargs):
    """
    生成单个基金产品各类指标(不包括月度/年度指标)汇总表[对普通基金或指增基金的绝对收益进行统计，而不是超额收益]

    Args:
        - netval_path (str): 净值数据路径，数据表第一列必须是日期，第二列必须是净值数据，且净值数据列名必须等于产品名
        - corp_name (str, optional): 私募管理人名称，如果没有输入该参数，默认是 "私募管理人"
        - start_date (date, optional): 可选参数，起始计算日期，可以不填，如果填写必须填 datetime.date 格式. 
    """
    netval_data = pd.read_excel(netval_path, index_col = 0)
    fund_name = netval_data.columns[0]
    start_date: date = kwargs.get("start_date", None)
    generate_word_indicator_tables(netval_data.iloc[:, 0], corp_name, start_date, fund_name = fund_name)

def multi_fund_indicator_tables(netval_path: str, **kwargs):
    """
    生成一只基金的报告，当然，在本函数体内写循环即可将它改为同时生成多只基金的报告。输出的文件在 output 文件夹下

    Args:
        - netval_path (str): 净值数据路径，数据表第一列必须是日期，第二列必须是净值数据，且净值数据列名必须等于产品名
        - corp_names (list[str]): 可选参数，私募管理人名称列表，如果没有输入该参数，默认是 "私募管理人"
        - start_dates (list[date]): 可选参数，起始计算日期列表，可以不填，如果填写必须填 datetime.date 格式. 

    """
    netval_data = pd.read_excel(netval_path, index_col = 0)
    funds_num: int = len(netval_data.columns)
    print(f"净值数据表中有{funds_num}只基金：", netval_data.columns)
    fund_names: list = netval_data.columns
    start_dates: list = kwargs.get("start_dates", [])
    corp_names: list = kwargs.get("corp_names", [])
    corp_names = [utils.CORP_DEFAULT_NAME if not elem else elem for elem in corp_names]
    start_dates += (funds_num - len(start_dates)) * [None]
    corp_names += (funds_num - len(corp_names)) * [utils.CORP_DEFAULT_NAME]
    for idx in tqdm(range(funds_num)):
        generate_word_indicator_tables(netval_data.iloc[:, idx], corp_names[idx], start_dates[idx], fund_name = fund_names[idx])