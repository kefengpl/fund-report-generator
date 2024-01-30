""" 此文件用于生成指增或者非指增基金的产品分析部分的WORD """
import os
from tqdm import tqdm
import numpy as np
import pandas as pd
from datetime import date

import fund
import enhanced_fund as ef
import word_handler as wh
import utils

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

def single_fund_indicator_tables(netval_path: str, corp_name: str = "私募管理人", **kwargs):
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

def multi_fund_report(netval_path: str, index_path: str, enhanced_fund: bool, **kwargs):
    """
    生成一只基金的报告，当然，在本函数体内写循环即可将它改为同时生成多只基金的报告。输出的文件在 output 文件夹下

    Args:
        - netval_path (str): 净值数据路径，数据表第一列必须是日期，第二列必须是净值数据，且净值数据列名必须等于产品名
        - index_path (str): 指数数据文件路径，数据表第一列必须是日期，后面的列可以传入多个指数数据(允许传入收盘价)，每列列名必须等于指数名
        - enhanced_fund (bool): 是否是指增基金
        - corp_names (list[str]): 可选参数，私募管理人名称列表，如果没有输入该参数，默认是 "私募管理人"
        - start_dates (list[date]): 可选参数，起始计算日期列表，可以不填，如果填写必须填 datetime.date 格式. 
        - add_indicators_tables (bool, optional): 可选参数，表示是否包含 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表


    """
    netval_data = pd.read_excel(netval_path, index_col = 0)
    index_data = pd.read_excel(index_path, index_col = 0)
    funds_num: int = len(netval_data.columns)
    print(f"净值数据表中有{funds_num}只基金：", netval_data.columns)
    fund_names: list = netval_data.columns
    index_name = index_data.columns[0] # 仅用于指增基金，用于给定标准的指数名称，便于后续处理
    start_dates: list = kwargs.get("start_dates", [])
    corp_names: list = kwargs.get("corp_names", [])
    start_dates += (funds_num - len(start_dates)) * [None]
    corp_names += (funds_num - len(corp_names)) * ["私募管理人"]
    add_indicators_tables: bool = kwargs.get("add_indicators_tables", False)
    for idx in tqdm(range(funds_num)):
        generate_report(netval_data.iloc[:, idx], index_data, enhanced_fund, corp_names[idx], start_dates[idx],
                        add_indicators_tables = add_indicators_tables, fund_name = fund_names[idx], index_name = index_name)

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
    start_dates += (funds_num - len(start_dates)) * [None]
    corp_names += (funds_num - len(corp_names)) * ["私募管理人"]
    for idx in tqdm(range(funds_num)):
        generate_word_indicator_tables(netval_data.iloc[:, idx], corp_names[idx], start_dates[idx], fund_name = fund_names[idx])

def generate_report(netval_data: pd.Series, index_data: pd.DataFrame, enhanced_fund: bool,
                    corp_name: str = "私募管理人", start_date: date = None, create_date: date = None, 
                    add_indicators_tables: bool = False, **kwargs):
    """
    生成单个基金产品报告的WORD。填入参数时注意参数类型。pd.Sries和pd.DataFrame是两种类型，需要区分。
    start_date是可选参数，表示开始计算的日期，如果是None，则会默认从传入数据的第一个有净值数据的日期开始计算

    Args:
        - netval_data (pd.Series): 净值数据，最左侧列需要是日期，读取数据时注意必须加：index_col = 0。
        - fund_name (str): 基金名称。
        - index_data (pd.DataFrame): 指数数据，最左侧列需要是日期，传入原始的收盘价即可，读取数据时注意必须加：index_col = 0。
        - corp_name (str): 该基金对应的私募管理人名称，可以不填。
        - enhanced_fund (bool): 是否是指增基金
        - start_date (date, optional): 起始计算日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - create_date (date, optional): 基金成立日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - add_indicators_tables (bool, optional): 可选参数，是否包含 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表
        - kwargs: 其它可选参数，需要个性化定制，目前支持的可选参数如下：\n
            ① analyze_text_start_year ，可选参数，它表示获得分析文本的年度收益时，从哪一年开始 \n
            ② history_table_start_year ，可选参数，它表示 历史收益数据表计算月度数据时，从哪一年开始
    """
    # PART0: 设置输出文件夹，如果有，就不管；如果没有，则创建 output 文件夹
    utils.create_output_folder()

    # PART0: 接收输入的参数
    fund_name = kwargs.get("fund_name")
    index_name = kwargs.get("index_name")
    analyze_text_start_year = kwargs.get("analyze_text_start_year", None)
    history_table_start_year = kwargs.get("history_table_start_year", None)
    this_fund = ef.EnhancedFund(fund_name, netval_data, index_data, index_name, start_date, create_date) \
                if enhanced_fund else fund.Fund(fund_name, netval_data, start_date, create_date)
    
    # PART1：获取生成word所需要的数据
    return_risk_table = this_fund.return_risk_table() # 获取表格“收益风险指标”所有单元格的数据
    history_return_table = property_method(enhanced_fund, "history_return_table", this_fund)(history_table_start_year) # 获取“历史收益分析”所有单元格的数据
    # 获取绘图数据，由于二者参数不一致，所以无法重载
    draw_data = this_fund.get_chart_data() if enhanced_fund else this_fund.get_chart_data(index_data)
    file_name = this_fund.export_chart_data(draw_data) # 导出绘图数据，并返回文件名
    file_abs_path = os.path.abspath("output/" + file_name) # 由于win32py只有绝对路径，所以把相对路径转为绝对路径
    # 获取分析文本那一段话
    analyze_text = property_method(enhanced_fund, "get_analyze_text", this_fund)(analyze_text_start_year) 
    footer_text = this_fund.get_footnote_text(corp_name) # 获取两个表格的脚注文本
    blank_fill = "超额" if enhanced_fund else ""

    # PART2：开始写入 WORD
    word_handler = wh.WordHandler(visible = False) 
    word_handler.set_page_layout() # 把 A4 纸横过来
    # 生成标题
    word_handler.add_text_content("1. " + this_fund.fund_name, "title")
    # 生成净值走势图和脚注
    word_handler.add_excel_chart(file_abs_path)
    word_handler.add_text_content("数据来源：" + corp_name + "，Wind", "footnote")
    word_handler.add_text_content("", "footnote")
    # 生成标题
    word_handler.add_text_content("1) 业绩分析", "title")
    word_handler.add_text_content("1.1) 收益走势", "title")
    # 生成产品分析文本
    word_handler.add_text_content(analyze_text)
    # 生成标题
    word_handler.add_text_content("1.2) 收益风险指标", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“收益风险指标”表格及其脚注
    word_handler.add_table(return_risk_table.shape[0], return_risk_table.shape[1], 
                           "weak_sep" if enhanced_fund else "sep", return_risk_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")
    # 生成标题
    word_handler.add_text_content(f"1.3) 历史{blank_fill}收益分析", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“历史收益分析”表格及其脚注
    word_handler.add_table(history_return_table.shape[0], history_return_table.shape[1], "first_row", history_return_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")

    if add_indicators_tables: # 添加补充的三张表格
        generate_word_indicator_tables(netval_data, corp_name, start_date, create_date, word_handler, this_fund, fund_name = fund_name)
        return # 上面的函数自己已经包含了保存并退出的功能
    
    # 保存文件并退出
    word_handler.close_and_save(this_fund.fund_name)
    # 生成并打印警告信息
    print_warning_messages(this_fund.fund_name, this_fund.get_first_netval_date())



def generate_word_indicator_tables(netval_data: pd.Series,  corp_name: str = "私募管理人", 
                                   start_date: date = None, create_date: date = None, 
                                   word_handler: wh.WordHandler = None, this_fund: fund.Fund = None, **kwargs):
    """
    生成单个基金产品各类指标(不包括月度/年度指标)汇总表[不包括指增基金]，滚动收益率分位数表，盈利概率表。
    填入参数时注意参数类型。pd.Sries和pd.DataFrame是两种类型，需要区分。
    start_date是可选参数，表示开始计算的日期，如果是None，则会默认从传入数据的第一个有净值数据的日期开始计算

    Args:
        - netval_data (pd.Series): 净值数据，最左侧列需要是日期，读取数据时注意必须加：index_col = 0。
        - corp_name (str): 该基金对应的私募管理人名称，可以不填。
        - start_date (date, optional): 起始计算日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - create_date (date, optional): 基金成立日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - word_handler (wh.WordHandler, optional):  如果传入了该参数并且合法，则会在该 word 里面追加写入内容，而不会新建一个 word 文档
        - this_fund (fund.Fund, optional): 基金计算对象，如果是空的话会新建一个，否则会沿用原来的对象。
        - fund_name (str): 基金名称(可选参数，在**kwargs中)。
    """
    # PART0: 设置输出文件夹，如果有，就不管；如果没有，则创建 output 文件夹
    utils.create_output_folder()

    # PART0: 接收输入的参数
    if not this_fund:
        fund_name = kwargs.get("fund_name")
        this_fund = fund.Fund(fund_name, netval_data, start_date, create_date)
        output_file_name: str = this_fund.fund_name + '_指标汇总及滚动收益统计'
    else:
        output_file_name: str = this_fund.fund_name
    
    # PART1：获取生成word所需要的数据[也就是三张表的数据]，如果是指增基金，则统计的是超额净值的滚动情况
    # 这里通过反射来调用合适的函数，表示是获取超额净值的数据还是普通基金净值的数据
    enhanced_fund: bool = isinstance(this_fund, ef.EnhancedFund)
    summary_indicators = property_method(enhanced_fund, "summary_indicators", this_fund)()
    all_recent_return = property_method(enhanced_fund, "all_recent_return", this_fund)()
    rolling_quantile_dataframe = property_method(enhanced_fund, "get_rolling_quantile_dataframe", this_fund)()
    earning_probability = property_method(enhanced_fund, "get_earning_probability", this_fund)()

    summary_indicator_table = np.r_[utils.dict_to_matrix(summary_indicators), 
                                    utils.dict_to_matrix(all_recent_return)]
    rolling_quantile_table = utils.df_to_matrix(rolling_quantile_dataframe)
    earning_probability_table = utils.df_to_matrix(earning_probability)
    footer_text = this_fund.get_footnote_text(corp_name) # 获取表格的脚注文本

    blank_fill: str = "" if not enhanced_fund else "超额" 

    series_list = ["1.4)", "1.5)", "1.6)"]
    # PART2：开始写入 WORD 
    if word_handler is None:
        series_list = ["1.", "2.", "3."] 
        word_handler = wh.WordHandler(visible = False)
        word_handler.set_page_layout()

    # 生成标题
    word_handler.add_text_content(f"{series_list[0]} {this_fund.fund_name}{blank_fill}关键指标汇总", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“收益风险指标”表格及其脚注
    word_handler.add_table(summary_indicator_table.shape[0], summary_indicator_table.shape[1], 
                           "row_sep", summary_indicator_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote") 

    # 生成标题
    word_handler.add_text_content(f"{series_list[1]} {this_fund.fund_name}{blank_fill}滚动收益率分布", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“收益风险指标”表格及其脚注
    word_handler.add_table(rolling_quantile_table.shape[0], rolling_quantile_table.shape[1], 
                           "first_row", rolling_quantile_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")

    # 生成标题
    word_handler.add_text_content(f"{series_list[2]} {this_fund.fund_name}{blank_fill}收益概率统计", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“收益风险指标”表格及其脚注
    word_handler.add_table(earning_probability_table.shape[0], earning_probability_table.shape[1], 
                           "first_row", earning_probability_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")

    # 保存文件并退出
    word_handler.close_and_save(output_file_name)
    # 打印警告信息
    print_warning_messages(this_fund.fund_name, this_fund.get_first_netval_date())

def property_method(enhanced_fund: bool, method_name: str, this_fund):
    """
    根据是否是指增基金调用合适的方法，返回的是方法对象。适用于一些需要调用 excess 相关方法的情况

    Args:
        enhanced_fund (True): 是否是指增基金
        method_name (str): 调用的方法名称
    """
    return getattr(this_fund.excess, method_name) if enhanced_fund else getattr(this_fund, method_name)

def print_warning_messages(fund_name: str, first_netval_date: date):
    """
    打印警告信息

    Args:
        - fund_name (str): _description_
        - first_netval_date (date): _description_
    """
    finish_message = f"----------------------------------{fund_name} 已经计算完成----------------------------------"
    # 生成并打印警告信息
    warning_message1 = f"警告信息1：年化收益率的起始计算日期是 {fund_name} 第一个有净值数据的日期 {first_netval_date} ，不是周报中的基金成立日期。" \
                       + "一般情况下计算结果与周报的年化数值十分接近。暂时不支持按照基金成立日期计算年化收益，这是因为①根据利息理论，代码中的算法没有问题；" \
                       + "②基于兼容性考虑的，因为代码支持用户指定起始计算日期 start_date，如果指定 start_date 后还使用基金成立日期计算年化收益，将得到完全错误的结果。" \
                       + "如果实在希望使用周报的起始日期计算年化，也是有一些解决方案的：将 annual_return 计算的起始日期 从 第一个净值日期 变更为 基金成立日期 即可"
   
    warning_message2 = "警告信息2：由于警告信息1，将使得夏普比、卡玛比、索提诺比的结果可能会有偏差，这是正常现象，不代表出现了错误，一般情况下偏差会很小。"
    # 打印警告信息
    print(finish_message)
    print(warning_message1)
    print(warning_message2)