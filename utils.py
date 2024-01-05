""" 此文件之作用在于：提供一些大家喜闻乐见的工具函数或者共用变量，或许会用于各个模块 """

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
    return file_name + time_str + _suffix

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