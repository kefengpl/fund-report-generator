""" 
此文件之作用在于将单个值或者可迭代类型的日期元素转化为 datetime.date
此文件的另一个作用在于匹配与年末或者月末最接近的日期，例如，为2016-12-31 寻找距离它最近的交易日
为 2023-02-28 (月末日期) 寻找最近的匹配日期
"""
from datetime import date
from datetime import datetime
import calendar

def scalar_to_date(input_date) -> date:
    """
    将输入格式转化为 date 格式，

    Args:
        - input_date (_type_): 输入的日期(标量)，可以是 str, datetime.datetime, datetime.date
        - str 支持 2015-01-02 或者 2015/1/2 两种类型

    Returns:
        date: date 格式的输入日期
    """
    if type(input_date) == str:
        separator = "/" if "/" in input_date else "-"
        _format = "%Y" + separator + "%m" + separator + "%d" 
        input_date = datetime.strptime(input_date, _format).date()
    try:
        input_date = input_date.date()
    except:
        raise ValueError("数据中的日期数据存在问题，程序无法将其转换为 datetime.date 类型")
    return input_date

def list_to_date(date_list):
    """
    将输入可迭代对象的元素格式转化为 date 格式

    Args:
        date_list (_type_): 任何可迭代对象
    """
    return list(map(lambda elem: scalar_to_date(elem), date_list))

def last_date_of_month(month: int, year: int):
    """
    返回每个月的最后一天

    Args:
        - month (int): _description_
        - year (int): _description_

    Returns:
        _type_: _description_
    """
    if month in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif month in [4, 6, 9, 11]:
        return 30
    elif month == 2:
        return 29 if calendar.isleap(year) else 28

def date_in_range(target_date: date, date_list: list) -> bool:
    """
    查找某个日期是否在给定列表范围内，即 >= 列表最小值; <= 列表最大值

    Args:
        - target_date (date): 要查询的目标日期
        - date_list (list[date]): 日期列表，它的列表中的格式必须是 date 类型

    Returns:
        bool: 如果 target_date 在范围内，就返回True；例如，某个列表最小日期是
        2017-01-01，最大日期是2023-12-23，那么 2016-12-23不在范围内，2017-01-03在范围内
    """
    return (target_date >= min(date_list)) and (target_date <= max(date_list)) 


def find_closest_date(target_date: date, date_list: list) -> date:
    """
    查找列表 search_list 中距离 target_date 最近的日期，主要用于在列表中匹配月末日期以及季度日期

    Args:
        - target_date (date): 要查询的目标日期
        - date_list (list[date]): 日期列表，它的列表中的格式必须是 date 类型，否则会有意想不到的错误

    Returns:
        date: 匹配到的最靠近日期，如果日期不在列表范围内，返回 None
    """
    if not date_in_range(target_date, date_list):
        return None
    date_diff = [abs((target_date - elem).days) for elem in date_list]
    return date_list[date_diff.index(min(date_diff))]