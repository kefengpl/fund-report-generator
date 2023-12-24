""" 此代码之作用在于逐步替代原有交互系统 """
import numpy as np
import pandas as pd
from datetime import date
from datetime import datetime

import date_handler

class Fund:
    def __init__(self, fund_name: str, create_time: date, net_val: pd.Series):
        """
        构造一个基金类，它存储了基金的净值数据，成立日期，基金名称

        Args:
            - fund_name (str): 基金名称
            - create_time (date): 基金成立日期，必须是 datetime.date 格式
            - net_val (pd.Series): 基金净值数据，index 是时间序列，datetime 或者 date 格式
            - basic_data (pd.DataFrame): 表示周报计算中的四列 "净值数据" "周度收益" "回撤" "标准化"。
              例如：沣京价值增强一期 沣京价值增强一期-收益率 沣京价值增强一期-回撤  沣京价值增强一期-标准化
              它是导出指标，可以自动计算。
        """
        if len(net_val) == 0:
            raise ValueError("你传入的参数没有任何数据，禁止构建此对象")
        # 日期统一为 datetime.date 格式，也就是有三个属性 year month day
        net_val.index = date_handler.list_to_date(net_val.index)
        create_time = date_handler.scalar_to_date(create_time)
        #net_val, create_time = self.pre_handle_date(net_val, create_time)
        self.net_val = self.interpolation(net_val)
        self.create_time = create_time 
        self.fund_name = fund_name
        self.basic_data = self.get_basic_data()
    
    def get_first_netval_date(self) -> date:
        """
        获取第一个净值日期

        Returns:
            date: 该基金有净值数据的首个日期
        """
        return self.net_val.first_valid_index()

    def interpolation(self, net_val: pd.Series):
        """
        进行插值填充数据

        Args:
            net_val (pd.Series): 基金净值数据，index 是时间序列，datetime.date 格式
        """
        net_val = net_val.apply(pd.to_numeric, errors = 'coerce')  # 能转数值就转数值，不能的话就转 Nan
        return net_val.interpolate(method='linear') # 线性插值：首部缺失不填充，中部缺失取线性，尾部缺失取尾数
    
    def get_basic_data(self, contain_standard: bool = True) -> pd.DataFrame:
        """
        Args:
            contain_standard (bool, optional): 可选参数，表示是否包含标准化那一列

        Returns:
            pd.DataFrame: 导出的4列(如果 contain_standard是 False，那就是 3 列)基础数据
        """
        basic_data = pd.DataFrame(self.net_val, columns = [self.fund_name])
        basic_data[self.fund_name + "-收益率"] = basic_data[self.fund_name].pct_change() # 计算周度收益率
        basic_data[self.fund_name + "-回撤"] = self.calculate_drawdown() # 计算回撤
        if contain_standard: # 计算标准化
            basic_data[self.fund_name + "-标准化"] = self.net_val / self.net_val[self.net_val.first_valid_index()] 
        return basic_data
    
    def calculate_drawdown(self) -> pd.Series:
        """
        计算回撤数据

        Returns:
            pd.Series: 计算出的回撤数据，格式 pd.Series，index 是 self.net_val 的 index，即日期索引
        """
        calculate_result = []
        cur_max = np.nan # 记录直到当前的最大值
        for elem in self.net_val:
            cur_max = max(-1 if pd.isna(cur_max) else cur_max, 
                          -1 if pd.isna(elem) else elem)
            if pd.isna(elem) or pd.isna(cur_max):
                calculate_result.append(np.nan)
            else:
                calculate_result.append(elem / cur_max - 1)
        # 细节处理，第一个有效净值对应的回撤数据依然是 0 
        calculate_result = pd.Series(calculate_result, index = self.net_val.index)
        calculate_result[self.net_val.first_valid_index()] = np.nan
        return calculate_result

def main():
    data = pd.read_excel("data/中金量化测试数据.xlsx", index_col = 0)
    name_list = data.columns
    date_list = list(data.loc["日期"])
    data = data.drop("日期", axis = 0)
    the_fund = Fund(name_list[0], date_list[0], data[name_list[0]])
    the_fund.basic_data.to_excel("output.xlsx")

if __name__ == "__main__":
    main()



