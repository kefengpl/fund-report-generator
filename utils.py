""" 此文件之作用在于：提供一些大家喜闻乐见的工具函数或者共用变量，或许会用于各个模块 """
import numpy as np
import pandas as pd
from datetime import datetime
import os

# 存放 RGB 三元组。这些颜色的RGB值有的来自于“每周策略观察 PPT”
color_dict = {
    "deep_red" : (192, 0, 0),
    "orange" : (255, 192, 0),
    "deep_grey" : (127, 127, 127),
    "shallow_red" : (255, 196, 200),
    "shallow_blue" : (124, 199, 255),
    "middle_blue" : (0, 104, 183),
    "deep_blue" : (0, 106, 105),
    "blue_green" : (0, 142, 140),
    "shallow_grey" : (191, 191, 191)
}

def not_pct_indicator(indicator_name: str) -> bool:
    """
    用于排除 夏普、索提诺、卡玛 这三个指标转化为百分比。它们只需要保留一位小数

    Args:
        indicator_name (str): 指标名称

    Returns:
        bool: 是不是 夏普、索提诺、卡玛、天数 三者之一，返回 True
    """
    check_list = ['夏普', '索提诺', '卡玛', 'Sortino', 'sortino', 'Calmar', 'Calmar', '天数']
    for elem in check_list:
        if elem in indicator_name:
            return False 
    return True

def convert_to_RGB(BGR_value: int) -> int:
    """ 通过位运算将 BGR 转为 RGB """
    blue = (BGR_value & 0xFF0000) >> 16
    green = (BGR_value & 0x00FF00) >> 8
    red = BGR_value & 0x0000FF
    return (red << 16) | (green << 8) | blue

def RGB_tuple_to_float(RGB_tuple: tuple) -> int:
    """ 将RGB三元组，例如(255, 196, 200)转化为 0xffc4c8 """
    return (RGB_tuple[2] << 16) | (RGB_tuple[1] << 8) | (RGB_tuple[0])

def decimal_to_pct(number: float) -> str:
    """ 小数转为百分比，保留一位小数 """
    return "{:.1%}".format(number)

def round_decimal(number: float) -> str:
    """ 小数保留1位小数 """
    return "{:.1f}".format(number)

def suitable_convert(number: float, indicator_name: str) -> str:
    """
    对字典中的value执行合理的转换。将合理的小数转为百分比并保留一位小数，将一些小数保留一位小数，将nan值转为 '-'

    Args:
        number (float): _description_
        indicator_name (str): value 对应的 key，也是指标名称

    Returns:
        str: 转换后的数值
    """
    if type(number) == str:
        return number
    if pd.isna(number):
        return '-'
    if not_pct_indicator(indicator_name):
        return decimal_to_pct(number)
    else:
        return round_decimal(number)

def dict_to_series(data: dict) -> pd.Series:
    """
    由于fund模块计算出的指标都是 dict 格式，某些时候也许难以处理，所以这里提供 dict -> pd.Series 

    Args:
        data (dict): fund模块相应函数的计算结果
    """
    return pd.Series(data.values(), data.keys())

def dict_to_dataframe(data: dict, columns: list = ["数值"]) -> pd.DataFrame:
    """
    由于fund模块计算出的指标都是 dict 格式，某些时候也许难以处理，所以这里提供 dict -> pd.DataFrame 

    Args:
        data (dict): fund模块相应函数的计算结果
    """
    return pd.DataFrame(dict_to_series(data), columns = columns)

def generate_filename(file_name: str, _suffix: str = ".xlsx") -> str:
    """
    产生文件名，会给文件后面加入一些数字码，防止频繁覆盖以往文件

    Args:
        - file_name (str): 希望的文件名称，不包含后缀名，例如 "导出数据"
        - _suffix (str, optional): 文件后缀名，默认是 ".xlsx".

    Returns:
        str: 添加数字码的文件名称
    """
    current_time: datetime = datetime.now()
    template_str = "{:02d}"
    time_str: str = str(current_time.year)[-2:] + template_str.format(current_time.month) + template_str.format(current_time.day) \
               + template_str.format(current_time.hour) + template_str.format(current_time.minute) + template_str.format(current_time.second)
    return file_name +  "_" + time_str + _suffix

def drop_suffix(_str: str, suffix: str = "标准化") -> str:
    """
    去掉字符串后缀，如 银河2号-标准化 ---> 银河2号
                   或 银河2号标准化  ---> 银河2号
    Args:
        _str (str): _description_
        suffix (str): _description_

    Returns:
        str: _description_
    """
    if _str.endswith("-" + suffix):
        return _str[:-(len(suffix) + 1)]
    if _str.endswith(suffix):
        return _str[:-len(suffix)]
    return _str

def create_output_folder():
    """ 创建 output 文件夹 """
    # 定义要创建的文件夹名称
    folder_name = "output"
    # 检查文件夹是否存在，如果不存在则创建它
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

def get_value(the_dict: dict):
    """ 它的作用是从具有单个键值对的字典中，比如 {"夏普比率" : 2.5} 中提取出数值 2.5 """
    if len(the_dict) != 1:
        raise ValueError(f"字典为空或者字典所含键值对的个数大于1，字典键值对个数{len(the_dict)}")
    return next(iter(the_dict.values()))

def df_to_matrix(df: pd.DataFrame) -> np.ndarray:
    """
    将输入的 df 数据框转化为 np.array 的矩阵(包括表头以及首列的列名称/行名称)
    便于直接在 word 中写入表格。此外，会将小数转化为保留1位小数的百分比

    Args:
        df (pd.DataFrame): 待转换的 dataframe

    Returns:
        np.ndarray: 转换为 np.array 后的矩阵
    """
    value_array = np.array(df)
    value_array = np.array([[decimal_to_pct(elem) if not pd.isna(elem) else '-' for elem in row] for row in value_array])
    colname_array = np.array([df.columns])
    indexname_array = np.array([df.index.name] + list(df.index))
    indexname_array = indexname_array.reshape(len(indexname_array), 1)
    return np.c_[indexname_array, np.r_[colname_array, value_array]]

def dict_to_matrix(indicators_dict: dict, column_number: int = 6) -> np.ndarray:
    """
    将一个字典数据转化为 np.array 的矩阵

    Args:
        - indicators_dict (dict): 一个字典，它汇总了各个指标及其数值，比如 {"夏普" : 0.9, "卡玛" : 0.75}
        - column_number (int): 这个数据表希望有几列？默认是6列。

    Returns:
        np.ndarray: 转换为 np.array 后的矩阵
    """
    indicators_names = list(indicators_dict.keys())
    indicators_values = list(indicators_dict.values())
    indicators_values = [suitable_convert(value, name) for (value, name) in zip(indicators_values, indicators_names)]
    sep_indicators_names = [indicators_names[idx : idx + column_number] 
                            for idx in range(0, len(indicators_names), column_number)]
    sep_indicators_values = [indicators_values[idx : idx + column_number] 
                            for idx in range(0, len(indicators_values), column_number)]
    sep_indicators_names[-1] += (column_number - len(sep_indicators_names[-1])) * [""]
    sep_indicators_values[-1] += (column_number - len(sep_indicators_values[-1])) * [""]
    final_matrix = []
    for names, values in zip(sep_indicators_names, sep_indicators_values):
        final_matrix.append(names)
        final_matrix.append(values)
    return np.array(final_matrix)