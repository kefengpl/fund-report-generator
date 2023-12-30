""" 此文件的作用在于，处理指增基金，它继承了 Fund 类，还包含一个 Fund 类 """
import pandas as pd
from datetime import date
# 引入 Fund 模块
from fund import Fund
import index_handler as ih

class EnhancedFund(Fund):
    def __init__(self, fund_name: str, net_val: pd.Series, index_data: pd.DataFrame, index_name: str, 
                 start_date: date = None, create_time: date = None):
        """
        此类用于处理指增数据。既继承了 fund 模块，内部又包含一个 fund 模块(用于计算超额部分的相关指标)

        Args:
            - fund_name (str): 基金名称
            - net_val (pd.Series): 基金净值数据，注意必须只包括净值数据，绝对不可以把成立日期那一行也包括进来
                                 索引index是时间序列，datetime 或者 date 格式
            - index_data (pd.DataFrame): 指数数据，最左侧列需要是日期，传入原始的收盘价即可，读取数据时注意必须加：index_col = 0。
                                      注意：只允许传入一个指数的数据，表示的是该指增基金对标的指数，且列名是该指数的名称
            - index_name (str): 指数的名称
            - start_date (date, optional): 希望从哪个日期开始计算，是人为指定的开始日期，其数值必须在传入数据的日期序列当中。
                                         默认值为 None，表示将从传入数据的首个有净值的日期开始计算. Defaults to None.
            - create_time (date, optional): 基金成立日期，必须是 datetime.date 格式，可以不填. Defaults to None.
        """
        if len(index_data.columns) >= 2:
            raise ValueError(index_data, "指数增强基金传入的指数数据只能包含一列，当前指数数据的列数：", len(index_data.columns))
        super().__init__(fund_name, net_val, start_date, create_time) # 调用父类构造函数
        self.index_data: pd.DataFrame = ih.IndexHandler(index_data, self.start_date, False).index_data # 指数收盘价预处理，但不标准化
        self.correct_index_dates() # 日期校准，修改 self.index_data，使得指数数据与基金数据的日期序列一致
        self.index_name = index_name # 指数的名称
        self.excess_return: pd.Series = self.get_excess_return() # 计算超额收益
        # NOTE 无论前面有没有指定 start_date，这里都不需要指定开始日期，因为 correct_index_dates() 已经将指数日期与基金日期对齐了
        self.excess: Fund = Fund(self.fund_name + "-超额", self.excess_return) # 用于计算超额收益的各项数据
    
    def correct_index_dates(self):
        """ 如果传入的指数数据的日期序列和基金净值的日期序列不一致，则校准指数数据日期序列，使得其与基金数据完全一致 """
        self.index_data = self.basic_data.merge(self.index_data, how = "left", 
                                                left_index = True, right_index = True)[[self.index_data.columns[0]]]
    
    def get_excess_return(self) -> pd.Series:
        """ 计算超额收益，第一个数值是1 """
        # 首先设置初始因子，即每个超额收益都要算的 (新基金净值/旧基金净值) / (新指数数据/旧指数数据)
        initial_factor = (self.net_val.pct_change() + 1) / (self.index_data.iloc[:, 0].pct_change() + 1) 
        # NOTE 根据周报显示，首个超额收益数据是1，需要手动添加，做法是：超额收益数据计算出来后首个有效数据的前面一个数据改为1
        initial_factor.iloc[list(initial_factor.index).index(initial_factor.first_valid_index()) - 1] = 1
        return initial_factor.cumprod() # 直接用累乘返回超额收益

    def get_chart_data(self) -> pd.Series:
        """ 重载方法，用于获取指增类基金的绘图数据 """
        merged_data = super().get_chart_data(self.index_data)
        merged_data.drop("最大回撤(右轴)", axis = 1)
        merged_data["超额收益最大回撤(右轴)"] = self.excess.basic_data[self.excess.get_column_name("回撤")]
        merged_data["超额收益"] = self.excess.basic_data[self.excess.get_column_name("标准化")]
        return merged_data
    
    
