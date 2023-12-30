""" 此文件用于处理指数标准化以及与基金数据的连接 """
import pandas as pd
from datetime import date
from datetime import datetime

import date_handler as dh

class IndexHandler:
    def __init__(self, index_data: pd.DataFrame, start_date: date = None, standalized: bool = True):
        """
        此类用于处理指数数据，包括标准化，以及产生并导出绘图所用的数据表

        Args:
            - index_data (pd.DataFrame): 可以是指数净值数据，也可以是收盘价数据
            - start_date (date, optional): 开始计算标准化的日期，一般可以和基金首个净值日期一致. 
                                         比如某个指数从 2023-07-01 开始计算，那么这天的净值就是1。
                                         默认值是None，表示以每个指数的第一个有效净值为基准进行标准化
            - standalized (bool): 是否对指数数据进行标准化处理？
        """
        self.index_data = index_data
        self.pre_handle_date()
        if standalized:
            self.standardlize_index(start_date) # 对数据进行标准化

    def pre_handle_date(self):
        """ 唯一作用是把该数据表的日期处理为 datetime.date """
        self.index_data.index = dh.list_to_date(self.index_data.index)

    def standardlize_index(self) -> pd.DataFrame:
        """ 获取指数数据的标准化数据(即把原始指数收盘数据处理为类似于单位净值的数据)，它会对原始数据进行修改 """
        for column in self.index_data.columns:
            self.index_data[column] = self.index_data[column] /  \
                                self.index_data[column][self.index_data[column].first_valid_index()]

    def standardlize_index(self, start_date: date) -> pd.DataFrame:
        """
        获取指数数据的标准化数据(即把原始指数收盘数据处理为类似于单位净值的数据)，它会对原始数据进行修改 \n
        注意：如果 start_date 为空，或出现 start_date 对应的收盘数据是空值，那么默认以第一个有效指数数据为基准进行计算

        Args:
            - index_data (pd.DataFrame): 输入的指数数据，可以包含多个指数，但这些指数数据的日期序列必须对齐
            - start_date (date): 开始计算标准化的日期，一般可以和基金首个净值日期一致，
                                比如某个指数从 2023-07-01 开始计算，那么这天的净值就是1

        Returns:
            pd.DataFrame: 标准化后的数据，并且从 start_date 开始截断
        """
        if start_date not in self.index_data.index:
            raise ValueError("开始日期start_date必须是指数数据所列日期之一")
        for column in self.index_data.columns:
            self.index_data[column] = self.index_data[column] /  \
                 self.index_data[column][start_date if (start_date is not None and not pd.isna(self.index_data[column][start_date])) \
                                         else self.index_data[column].first_valid_index()]
        return self.index_data[self.index_data.index >= start_date]