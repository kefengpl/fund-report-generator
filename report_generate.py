""" 此文件用于生成指增或者非指增基金的产品分析部分的WORD """
import os
import pandas as pd
from datetime import date

import fund
import word_handler as wh

def generate_report(netval_data: pd.Series, fund_name: str, index_data: pd.DataFrame, corp_name: str = "私募管理人",
                    start_date: date = None, create_date: date = None, **kwargs):
    """
    生成单个基金产品报告的WORD。填入参数时注意参数类型。pd.Sries和pd.DataFrame是两种类型，需要区分。
    start_date是可选参数，表示开始计算的日期，如果是None，则会默认从传入数据的第一个有净值数据的日期开始计算

    Args:
        - netval_data (pd.Series): 净值数据，最左侧列需要是日期，读取数据时注意必须加：index_col = 0。
        - fund_name (str): 基金名称。
        - index_data (pd.DataFrame): 指数数据，最左侧列需要是日期，传入原始的收盘价即可，读取数据时注意必须加：index_col = 0。
        - corp_name (str): 该基金对应的私募管理人名称，可以不填。
        - start_date (date, optional): 起始计算日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - create_date (date, optional): 基金成立日期，可以不填，如果填写必须填 datetime.date 格式. Defaults to None.
        - kwargs: 其它可选参数，需要个性化定制：① 可选参数： analyze_text_start_year ，它表示获得分析文本的年度收益时，从哪一年开始
    """
    this_fund = fund.Fund(fund_name, netval_data, start_date, create_date)
    # PART1：获取生成word所需要的数据
    return_risk_table = this_fund.return_risk_table() # 获取表格“收益风险指标”所有单元格的数据
    history_return_table = this_fund.history_return_table(2020) # 获取“历史收益分析”所有单元格的数据
    draw_data = this_fund.get_chart_data(index_data) # 获取绘图数据，
    file_name = this_fund.export_chart_data(draw_data) # 导出绘图数据，并返回文件名
    file_abs_path = os.path.abspath("output/" + file_name) # 由于win32py只有绝对路径，所以把相对路径转为绝对路径
    analyze_text = this_fund.get_analyze_text(kwargs.get("analyze_text_start_year", None)) # 获取分析文本那一段话
    footer_text = this_fund.get_footnote_text(corp_name) # 获取两个表格的脚注文本

    # PART2：开始写入 WORD
    word_handler = wh.WordHandler(visible = False)
    word_handler.set_page_layout() # 把 A4 纸横过来
    # 生成标题
    word_handler.add_text_content("1. " + this_fund.fund_name, "title")
    # 生成净值走势图和脚注
    word_handler.add_excel_chart(file_abs_path)
    word_handler.add_text_content("数据来源：" + corp_name + "，Wind; ", "footnote")
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
    word_handler.add_table(return_risk_table.shape[0], return_risk_table.shape[1], "sep", return_risk_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")
    # 生成标题
    word_handler.add_text_content("1.3) 历史收益分析", "title")
    word_handler.add_text_content("", "footnote")
    # 生成“历史收益分析”表格及其脚注
    word_handler.add_table(history_return_table.shape[0], history_return_table.shape[1], "first_row", history_return_table)
    word_handler.add_text_content(footer_text, "footnote")
    word_handler.add_text_content("", "footnote")
    # 保存文件并退出
    word_handler.close_and_save(this_fund.fund_name)