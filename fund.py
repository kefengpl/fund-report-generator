""" 
此代码之作用在于根据净值数据进行指标计算，暂时与原来500行的交互系统独立运行
在本模块更新完毕后，就可以删除 interaction_sys.py
"""
import math
import numpy as np
import pandas as pd
from datetime import date
from datetime import datetime

import date_handler as dh

class Fund:
    def __init__(self, fund_name: str, create_time: date, net_val: pd.Series):
        """
        构造一个基金类，它存储了基金的净值数据，成立日期，基金名称

        Args:
            - fund_name (str): 基金名称
            - create_time (date): 基金成立日期，必须是 datetime.date 格式
            - net_val (pd.Series): 基金净值数据，注意必须只包括净值数据，绝对不可以把成立日期那一行也包括进来
                                   index 是时间序列，datetime 或者 date 格式
            - basic_data (pd.DataFrame): 表示周报计算中的四列 "净值数据" "周度收益" "回撤" "标准化"。
              例如：沣京价值增强一期 沣京价值增强一期-收益率 沣京价值增强一期-回撤  沣京价值增强一期-标准化
              它是导出指标，可以自动计算。
        """
        if len(net_val) == 0:
            raise ValueError("你传入的参数没有任何数据，禁止构建此对象")
        # 日期统一为 datetime.date 格式，也就是有三个属性 year month day
        net_val.index = dh.list_to_date(net_val.index)
        create_time = dh.scalar_to_date(create_time)
        self.net_val = self.interpolation(net_val)
        self.create_time = create_time 
        self.fund_name = fund_name
        self.basic_data = self.get_basic_data()
        self.date_list = self.basic_data.index # 获得日期列表

    def get_column_name(self, search_name: str = None) -> str:
        """
        获取合适的列名。例如，当 self.basic_data 列名是
        [沣京价值增强一期 沣京价值增强一期-收益率 沣京价值增强一期-回撤  沣京价值增强一期-标准化]
        这四列，输入 "收益" 或者 "收益率" 得到 列名 "沣京价值增强一期-收益率";
        输入 "回撤" 得到 列名 "沣京价值增强一期-回撤"; 输入标准化，得到"沣京价值增强一期-标准化"
        如果该参数是 None，那么返回第一列列名 "沣京价值增强一期"

        Args:
            search_name (str): 输入 "收益""收益率"(这二者得到结果一致)  "回撤" "标准化" 以得到该对象的数据表中对应的列名
                               也可以输入 None，默认获取数据表 basic_data 第一列的列名

        Returns:
            str: 匹配到的数据表 basic_data 中的列名。如果匹配不到，代码自然会报错。
        """
        if not search_name:
            return self.basic_data.columns[0]
        return [column for column in self.basic_data.columns if (search_name in column.lstrip(self.fund_name))][0]
    
    def get_first_netval_date(self) -> date:
        """
        获取第一个净值日期

        Returns:
            date: 该基金有净值数据的首个日期
        """
        return self.net_val.first_valid_index()
    
    def get_first_date(self) -> date:
        """ 获取该基金净值数据表中的第一个日期，不是该基金第一个有净值数据的日期
        比如某些周报中的第一个日期是 2007/9/6 """
        return self.net_val.index[0]
    
    def get_last_date(self) -> date:
        """ 获取数据中的最新净值日期 """
        return self.net_val.index[-1]

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
    
    def get_proper_end_date(self, raw_end_date: date) -> date:
        """
        由于某月或者某年没有过完，2024/1/31 这种日期有可能超出数据表的范围，从而无法匹配与之最接近的日期，
        因此，设立此模块，对于一般的日期，比如 2022/12/31，没有超出净值数据事件范围，就正常匹配最接近的日期即可
        而对于最近一个月或者最新年份，最后一天设置为净值数据表的最后一个日期，例如，数据表最后一个日期是 2023-11-23
        那么在计算2023年11月的月度收益率时，我们会把该月结束日期改为 2023-11-23
        """
        return self.get_last_date() if not dh.date_in_range(raw_end_date, self.date_list) \
                                        else dh.find_closest_date(raw_end_date, self.date_list)

    def one_year_return(self, year: int) -> float:
        """ 根据净值计算某一年的收益率，注意：算法是尽量匹配年末值，例如2017年理论值是 2017/12/31 净值 除以 2016/12/31 净值 - 1，
        该方法会匹配与 2017/12/31 和 2016/12/31 最接近的日期，并寻找它们的净值数据,计算失败就返回 np.nan """
        try:
            raw_start_date = date(year - 1, 12, 31)
            raw_end_date = date(year, 12, 31)
            start_date = dh.find_closest_date(raw_start_date, self.date_list)
            end_date = self.get_proper_end_date(raw_end_date)
            return self.net_val[end_date] / self.net_val[start_date] - 1
        except:
            return np.nan
    
    def one_month_return(self, year: int, month: int) -> float:
        """
        根据净值数据计算给定年月的月度收益率
        例如，当人类希望计算 2023年4月的收益时，理论上，他应该找到 2023/4/30 的数值和 2023/3/31 的数值
        ，但是，现实很残酷，由于净值数据是周度数据，他往往需要在净值数据中分别找到距离 2023/4/30 最近的日期，
        以及 距离 2023/3/31 最近的日期。本函数就是在查找这些近似日期，找到对应的净值数据，从而得到收益率 
        如果数据不全或缺失，会直接返回 np.nan

        - 如果下面这俩参数都是空值，该函数将从输入数据的首个月份开始计算，比如从 2007年9月开始计算，当然，这样会
        产生很多空值。
        - 如果下面这俩参数中，year 不空，month 空，

        Args:
            - year (int): 输入 int 类型表示的年份值，比如2023，默认是 None
            - month (int): 输入 int 类型表示的月份值，比如12，默认是 None

        Returns:
            float: 计算出的该月的收益
        """
        try:
            # NOTE 一个困难的地方在于：每年 1 月份的特殊处理
            raw_start_year = year if month != 1 else year - 1
            raw_start_month = month - 1 if month != 1 else 12
            raw_start_date = date(raw_start_year, raw_start_month, 
                                  dh.last_date_of_month(raw_start_month, raw_start_year))
            raw_end_date = date(year, month, dh.last_date_of_month(month, year))
            start_date = dh.find_closest_date(raw_start_date, self.date_list)
            end_date = self.get_proper_end_date(raw_end_date)
            return self.net_val[end_date] / self.net_val[start_date] - 1
        except:
            return np.nan
    
    def all_year_return(self, start_year: int = None) -> dict:
        """
        获取所有年份的收益率，如果前面的年份收益数据无法计算，那么该年度得到 np.nan
        NOTE BUG 这种算法使得基金成立那年的数据是无法计算的，因为第一年数据不全，有待修复。

        Args:
            start_year (int, optional): 年度收益率计算起始值，默认是None，此时会从日期序列的第一个日期所在年份开始计算

        Returns:
            dict: 所有范围内年份的收益率数据
        """
        # 设定计算开始年份
        start_year: int = self.get_first_date().year if start_year is None else start_year
        # 设定计算结束年份
        final_year: int = self.get_last_date().year
        calculate_result = {}
        for year in range(start_year, final_year + 1):
            calculate_result[str(year) + "年"] = self.one_year_return(year)
        return calculate_result
    
    def all_month_return(self, start_point: tuple = (None, None)) -> dict:
        """
        获取所有月份的收益率，如果前面月份的收益率无法计算，那么该月份就得到 np.nan

        Args:
            start_point (tuple, optional): 一个元组，第一个括号内写年份，第二个括号内写月份
            你可以设置从何年何月开始计算，默认是(None, None).如果你只设置了年份，那么会从该年1月开始计算；
            如果你只设置了月份，那么该函数会报错.

        Returns:
            dict: 所有月度的收益率数据
        """
        if (not start_point[0]) and start_point[1]:
            raise ValueError("不允许只输入月份而不输入年份")
        start_year = None
        start_month = None
        if start_point[0] and start_point[1]:
            start_year, start_month = start_point[0], start_point[1]
        elif start_point[0] and not(start_point[1]):
            start_year, start_month = start_point[0], 1
        else:
            start_year, start_month = self.get_first_date().year, self.get_first_date().month
        
        final_year, final_month = self.get_last_date().year, self.get_last_date().month
        calculation_list = self.generate_month_list(start_year, start_month, final_year, final_month)
        calculate_result = {}
        for year, month in calculation_list:
            calculate_result[str(year) + "年" + str(month) + "月"] = self.one_month_return(year, month)
        return calculate_result
    
    def generate_month_list(self, start_year: int, start_month: int, 
                            end_year: int, end_month: int) -> list:
        """ 此函数用于产生从 开始年月 到 结束年月 的所有月份列表，其中，某年某月以元组表示
        此函数是 all_month_return 的辅助函数 """
        result = []
        for month in range(start_month, 12 + 1):
            result.append((start_year, month))
        for year in range(start_year + 1, end_year):
            for month in range(1, 12 + 1):
                result.append((year, month))     
        for month in range(1, end_month + 1):
            result.append((end_year, month))
        return result   
    
    def cumulative_return(self) -> dict:
        """ 计算累计收益，就是 最后的净值数据 / 首个净值数据 - 1 """
        return {"累计收益率" : self.net_val[self.get_last_date()] / self.net_val[self.get_first_netval_date()] - 1}
    
    def annual_return(self) -> dict:
        """ 计算年化收益，注意：这里的计算模式是 (1 + 累计收益) ^ (365 / (最新日期 - 首个净值日期))  - 1 \n
        NOTE BUG 注意！！时间那里不是 (最新日期 - 成立日期) ，故可能与周报数据有出入！！！  """
        annual_return = pow(1 + next(iter(self.cumulative_return().values())), 
                            365 / (self.get_last_date() - self.get_first_netval_date()).days) - 1
        return {"年化收益率" : annual_return}
    
    def max_drawdown(self) -> dict:
        """ 获取最大回撤 """
        # 首先需要获取回撤那一列的列名，找不到就报错
        column_name = self.get_column_name("回撤")
        return {"最大回撤" : self.basic_data[column_name].min()}
    
    def annual_volatility(self) -> dict:
        """ 获取年化波动率：NOTE BUG 注意：这里是直接计算的标准差，故可能与周报中的数据有出入 """
        # 首先需要获取收益率那一列的列名，找不到就报错
        column_name = self.get_column_name("收益")
        return {"年化波动率" : self.basic_data[column_name].std() * math.sqrt(52)}
    
    def sharpe_ratio(self, risk_free_rate: int = 0.015) -> dict:
        """ 获取夏普比率，无风险利率默认是 1.5% """
        return {"夏普比率" : (next(iter(self.annual_return().values())) - risk_free_rate) / 
                next(iter(self.annual_volatility().values()))}
    
    def weekly_win_rate(self) -> dict:
        """ 获取周胜率 = 大于0的周度收益 / 所有有效的收益率数值个数 """
        # 首先需要获取收益率那一列的列名，找不到就报错
        column_name = self.get_column_name("收益")
        weekly_return = self.basic_data[column_name]
        return {"周胜率" : len(weekly_return[weekly_return > 0]) / (~weekly_return.isna()).sum()}
    
    def summary_indicators(self) -> dict: 
        """ 汇总除了 年度收益 和 月度收益 之外的所有指标 """
        return {**self.cumulative_return(), **self.annual_return(), **self.max_drawdown(),
                **self.annual_volatility(), **self.sharpe_ratio(), **self.weekly_win_rate()}
 
def main():
    data = pd.read_excel("data/中金量化测试数据.xlsx", index_col = 0)
    # NOTE 必须把成立日期和净值数据分开，否则
    name_list = data.columns
    date_list = list(data.loc["日期"])
    data = data.drop("日期", axis = 0)
    the_fund = Fund(name_list[0], date_list[0], data[name_list[0]])
    #the_fund.basic_data.to_excel("output.xlsx")
    print(the_fund.all_month_return())

if __name__ == "__main__":
    main()



