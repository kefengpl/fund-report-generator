""" 此文件用于生成指增或者非指增基金的产品分析部分的WORD """
import os
import pandas as pd
from datetime import date

import fund
import enhanced_fund as ef
import word_handler as wh
import utils

def property_method(enhanced_fund: bool, method_name: str, this_fund):
    """
    根据是否是指增基金调用合适的方法，返回的是方法对象。适用于一些需要调用 excess 相关方法的情况

    Args:
        enhanced_fund (True): 是否是指增基金
        method_name (str): 调用的方法名称
    """
    return getattr(this_fund.excess, method_name) if enhanced_fund else getattr(this_fund, method_name)


def generate_report(netval_data: pd.Series, index_data: pd.DataFrame, enhanced_fund: bool,
                    corp_name: str = "私募管理人", start_date: date = None, create_date: date = None, **kwargs):
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
    # 保存文件并退出
    word_handler.close_and_save(this_fund.fund_name)

    # 生成并打印警告信息
    warning_message1 = f"警告信息1：年化收益率的起始计算日期是 {this_fund.fund_name} 第一个有净值数据的日期 {this_fund.get_first_netval_date()} ，不是周报中的基金成立日期。" \
                       + "一般情况下计算结果与周报的年化数值十分接近。暂时不支持按照基金成立日期计算年化收益，这是因为①根据利息理论，代码中的算法没有问题；" \
                       + "②基于兼容性考虑的，因为代码支持用户指定起始计算日期 start_date，如果指定 start_date 后还使用基金成立日期计算年化收益，将得到完全错误的结果。" \
                       + "如果实在希望使用周报的起始日期计算年化，也是有一些解决方案的：将 annual_return 计算的起始日期 从 第一个净值日期 变更为 基金成立日期 即可"
    warning_message2 = "警告信息2：年化波动率的计算是直接使用统计学中的标准差公式计算的，与周报中的计算可能有出入" \
                       + "如果需要按照周报计算规则修改，请对 fund.py 中的 annual_volatility 函数进行修改"
    warning_message3 = "警告信息3：由于前两个警告信息，将使得夏普比、卡玛比、索提诺比的结果可能会有偏差，这是正常现象，不代表出现了错误，一般情况下偏差会很小。"
    # 打印警告信息
    print(warning_message1)
    # print(warning_message2)
    print(warning_message3)