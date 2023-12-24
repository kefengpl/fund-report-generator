""" 此文件用于处理指数标准化以及与基金数据的连接 """
import pandas as pd
from datetime import date
from datetime import datetime

import date_handler as dh

def pre_handle_date(index_data: pd.DataFrame) -> pd.DataFrame:
    """ 唯一作用是把该数据表的日期处理为 datetime.date """
    index_data.index = dh.list_to_date(index_data.index)
    return index_data

def standardlize_index(index_data: pd.DataFrame) -> pd.DataFrame:
    """ 获取指数数据的标准化数据(即把原始指数收盘数据处理为类似于单位净值的数据)，它会对原始数据进行修改 """
    for column in index_data.columns:
        index_data[column] = index_data[column] /  \
                             index_data[column][index_data[column].first_valid_index()]
    return index_data

def standardlize_index(index_data: pd.DataFrame, start_date: date) -> pd.DataFrame:
    """
    获取指数数据的标准化数据(即把原始指数收盘数据处理为类似于单位净值的数据)，它会对原始数据进行修改

    Args:
        - index_data (pd.DataFrame): 输入的指数数据
        - start_date (date): 开始计算标准化的日期，一般可以和基金首个净值日期一致，
                             比如某个指数从 2023-07-01 开始计算，那么这天的净值就是1

    Returns:
        pd.DataFrame: 标准化后的数据，并且从 start_date 开始截断
    """
    if start_date not in index_data.index:
        raise ValueError("开始日期start_date必须是指数数据所列日期之一")
    for column in index_data.columns:
        index_data[column] = index_data[column] /  \
                             index_data[column][start_date]
    return index_data[index_data.index >= start_date]