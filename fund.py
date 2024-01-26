""" 
此代码之作用在于根据净值数据进行指标计算
"""
import math
import numpy as np
import pandas as pd
from datetime import date
from datetime import datetime
from dateutil.relativedelta import relativedelta

import date_handler as dh
import index_handler as ih
import utils

class Fund:
    def __init__(self, fund_name: str, net_val: pd.Series, start_date: date = None, create_time: date = None):
        """
        构造一个基金类，它存储了基金的净值数据，成立日期，基金名称

        Args:
            - fund_name (str): 基金名称
            - create_time (date): 基金成立日期，必须是 datetime.date 格式，可以不填
            - net_val (pd.Series): 基金净值数据，注意必须只包括净值数据，绝对不可以把成立日期那一行也包括进来
                                   index 是时间序列，datetime 或者 date 格式
            - start_date(date): 希望从哪个日期开始计算，是人为指定的开始日期，其数值必须在传入数据的日期序列当中
                                默认值为 None，表示将从传入数据的首个有净值的日期开始计算
            - basic_data (pd.DataFrame): 表示周报计算中的四列 "净值数据" "周度收益" "回撤" "标准化"。
              例如：沣京价值增强一期 沣京价值增强一期-收益率 沣京价值增强一期-回撤  沣京价值增强一期-标准化
              它是导出指标，可以自动计算。
        """
        if len(net_val) == 0:
            raise ValueError("你传入的参数没有任何数据，禁止构建此对象")
        # 日期统一为 datetime.date 格式，也就是有三个属性 year month day
        net_val.index = dh.list_to_date(net_val.index)
        create_time = dh.scalar_to_date(create_time) if create_time is not None else create_time
        self.net_val = self.interpolation(net_val)
        self.create_time = create_time # NOTE 该变量似乎没有在后面的代码中使用
        self.fund_name = fund_name
        if start_date is not None:
            if start_date not in self.net_val.index:
                raise ValueError(start_date, "开始日期必须在传入数据的日期序列里")
            self.net_val = self.net_val[self.net_val.index >= start_date] # 手动设置起始日期后会截取净值数据
        self.basic_data = self.get_basic_data()
        self.date_list = self.basic_data.index # 获得日期列表
        self.rolling_return_data = self.get_rolling_return_data()  # 滚动收益数据表

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
        获取第一个净值日期。
        显然，给定 start_date 的情况下，由于存在数据截断，所以该函数依然适用，即它也能取出 
        从 start_date 开始的第一个有净值数据的日期

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
    
    def err_check_start_year_ge_first(self, start_year: int):
        """
        检查后面一些函数的起始年份是否在净值数据表的首年及之后。
        (允许前面几年净值数据是空值) ge 代表 greater equal (>=)
        比如，如果传入周报的某些数据，从2007年开始。那么 start_year >= 2007 即可

        Args:
            start_year (int): 开始年份
        """
        if start_year is not None and start_year < self.get_first_date().year:
            raise ValueError("错误值：", start_year, " 起始年份不得早于：", self.get_first_date().year)
        
    def err_check_start_year_ge_first_netval(self, start_year: int):
        """
        检查后面一些函数的起始年份是否在净值数据不为空的首年及之后。
        (允许前面几年净值数据是空值) ge 代表 greater equal (>=)
        比如，如果传入周报的某些数据，虽然从2007年开始，但是某基金产品净值数据从2018年开始，那么
        需要 start_year >= 2018

        Args:
            start_year (int): 开始年份
        """
        if start_year is not None and start_year < self.get_first_netval_date().year:
            raise ValueError("错误值：", start_year, " 起始年份不得早于：", self.get_first_netval_date().year)
        
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
    
    def get_rolling_return_data(self) -> pd.DataFrame:
        """ 获得半年、一年、二年、三年、五年的滚动收益，列名恰好是 半年、一年、二年、三年、五年 """
        rolling_return_data = pd.DataFrame()
        window_size = 50 # 同一计算窗口，一年的窗口期是50
        rolling_return_data["半年"] = self.net_val.pct_change(periods = window_size // 2)
        rolling_return_data["一年"] = self.net_val.pct_change(periods = window_size)
        rolling_return_data["二年"] = self.net_val.pct_change(periods = window_size * 2)
        rolling_return_data["三年"] = self.net_val.pct_change(periods = window_size * 3)
        rolling_return_data["五年"] = self.net_val.pct_change(periods = window_size * 5)
        return rolling_return_data
    
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
    
    def get_proper_start_date(self, raw_start_date: date, year: int, month: int = None) -> date:
        """
        由于某个基金可能是从一年的中间开始有净值的，raw_start_date 一般是上一年的12月31日，
        则此时基金成立的首个年度或者首个月度，开始日期应设为首个净值日期.
        如果 year 参数为 None，直接报错。如果有 year 但没有 month 参数，则仅匹配净值开始的年份。
        如果有 year 且 month 参数，则匹配净值开始的月份。
        """
        if not year:
            raise ValueError("年份参数 year 不能是空值")
        check_year: bool = (year == self.get_first_netval_date().year)
        check_month: bool = (month == self.get_first_netval_date().month) if month is not None else True
        return  self.get_first_netval_date() if check_year and check_month  \
                else dh.find_closest_date(raw_start_date, self.date_list) 


    def one_year_return(self, year: int) -> float:
        """ 根据净值计算某一年的收益率，注意：算法是尽量匹配年末值，例如2017年理论值是 2017/12/31 净值 除以 2016/12/31 净值 - 1；
        该方法会匹配与 2017/12/31 和 2016/12/31 最接近的日期，并寻找它们的净值数据,计算失败就返回 np.nan；
        NOTE 一些更新：基金净值日期的第一年也允许计算收益率，但是并不是全年的收益率，需要注意 """
        try:
            raw_start_date = date(year - 1, 12, 31)
            raw_end_date = date(year, 12, 31)
            start_date = self.get_proper_start_date(raw_start_date, year)
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
            start_date = self.get_proper_start_date(raw_start_date, year, month)
            end_date = self.get_proper_end_date(raw_end_date)
            return self.net_val[end_date] / self.net_val[start_date] - 1
        except:
            return np.nan
    
    def all_year_return(self, start_year: int = None) -> dict:
        """
        获取所有年份的收益率，如果前面的年份收益数据无法计算，那么该年度得到 np.nan
        NOTE 一些更新：基金净值日期的第一年允许计算收益率，但是并不是全年的收益率，需要注意
        NOTE 基金净值最后一年允许计算收益率，但是表示 当年以来的收益率，并不是全年的收益率

        Args:
            start_year (int, optional): 年度收益率计算起始值，默认是None，此时会从日期序列的第一个日期所在年份开始计算

        Returns:
            dict: 所有范围内年份的收益率数据
        """
        self.err_check_start_year_ge_first(start_year)
        
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
        - NOTE 一些更新：基金净值日期的第一年允许计算收益率，但是并不是全年的收益率，需要注意
        - NOTE 基金净值最后一年允许计算收益率，但是表示 当年以来的收益率，并不是全年的收益率

        Args:
            start_point (tuple, optional): 一个元组，第一个括号内写年份，第二个括号内写月份
            你可以设置从何年何月开始计算，默认是(None, None).如果你只设置了年份，那么会从该年1月开始计算；
            如果你只设置了月份，那么该函数会报错. \n

        Returns:
            dict: 所有月度的收益率数据
        """
        if (not start_point[0]) and start_point[1]:
            raise ValueError("不允许只输入月份而不输入年份")
        self.err_check_start_year_ge_first(start_point[0])
        if start_point[0] is not None and start_point[0] == self.get_first_date().year and start_point[1] < self.get_first_date().month:
            raise ValueError("错误值：", start_point, "起始年份的月份不得早于净值数据表中最早日期的月份：", self.get_first_date().month)
        
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
    
    def recent_month_return(self, month: int) -> float:
        """
        以输入数据的最新净值日期为基准，计算最近 month 个月的收益率
        比如最新日期是2023-10-20， month = 3，代码会自动匹配日期序列中最接近 2023-7-20 的日期

        Args:
            month (int): 表示的含义是 "近month月"，
                         比如 month = 6 表示近6个月，month = 12 表示近一年，month = 24 表示近 2 年 

        Returns:
            float: 近 month 月的收益率数据
        """
        try:
            final_date: date = self.get_last_date()
            begin_date: date = dh.find_closest_date(final_date - relativedelta(months = month), self.date_list)
            return self.net_val[final_date] / self.net_val[begin_date] - 1
        except:
            np.nan
    
    def all_recent_return(self) -> dict:
        """ 获取 近一月、近三月、近六月、近一年、近两年、近三年 的收益率 """
        recent_months: list = [1, 3, 6, 12, 24, 36]
        indicator_list: list = ["近一月", "近三月", "近六月", "近一年", "近两年", "近三年"]
        result_dict: dict = {}
        for month, indicator_name in zip(recent_months, indicator_list):
            result_dict[indicator_name] = self.recent_month_return(month)
        return result_dict
    
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
        """ 获取历史最大回撤 """
        # 首先需要获取回撤那一列的列名，找不到就报错
        column_name = self.get_column_name("回撤")
        return {"最大回撤" : self.basic_data[column_name].min()}
    
    def max_drawdown_of_recent_year(self) -> dict:
        """ 获取最近一年最大回撤，不足一年的情况下，该函数相当于获取历史最大回撤
            计算方式：比如最新日期 2023-10-20，函数会寻找最接近 2022-10-20 的日期，并计算 [2022-10-20, 2023-10-20] 闭区间内的最大回撤 """
        one_year_ago: date = dh.find_closest_date(self.get_last_date() - relativedelta(years = 1), self.date_list)
        column_name = self.get_column_name("回撤")
        indicator_name = "过去一年最大回撤"
        if not one_year_ago:
            return {indicator_name : self.max_drawdown()["最大回撤"]}
        return {indicator_name : self.basic_data[column_name][one_year_ago:].min()}
    
    def max_weekly_drawdown(self) -> dict:
        """ 计算最大周度回撤 """
        indicator_name = "最大周度回撤"
        column_name = self.get_column_name("收益率")
        return {indicator_name : self.basic_data[column_name].min()}
    
    def this_week_return(self) -> dict:
        """ 获取最新一周的收益率 """
        indicator_name = "本周收益率"
        column_name = self.get_column_name("收益率")
        return {indicator_name : self.basic_data[column_name][self.get_last_date()]}
    
    def annual_volatility(self) -> dict:
        """ 获取年化波动率：NOTE BUG 注意：这里是直接计算的标准差，故可能与周报中的数据有出入 """
        # 首先需要获取收益率那一列的列名，找不到就报错
        column_name = self.get_column_name("收益")
        return {"年化波动率" : self.basic_data[column_name].std() * math.sqrt(52)}
    
    def sharpe_ratio(self, risk_free_rate: float = 0.015) -> dict:
        """ 获取夏普比率，无风险利率默认是 1.5% """
        return {"夏普比率" : (utils.get_value(self.annual_return()) - risk_free_rate) / 
                utils.get_value(self.annual_volatility())}
    
    def weekly_win_rate(self) -> dict:
        """ 获取周胜率 = 大于0的周度收益 / 所有有效的收益率数值个数 """
        # 首先需要获取收益率那一列的列名，找不到就报错
        column_name = self.get_column_name("收益")
        weekly_return = self.basic_data[column_name]
        return {"周胜率" : len(weekly_return[weekly_return > 0]) / (~weekly_return.isna()).sum()}
    
    def decline_std(self) -> dict:
        """ 计算下行标准差，公式与周报计算一致。求解时只求平方和而不减去均值 """
        return_data = self.basic_data[self.get_column_name("收益率")]
        return {"下行标准差" : math.sqrt((return_data[return_data < 0] ** 2).sum() / (return_data.count() - 1))}
    
    def decline_std_annualize(self) -> dict:
        """ 年化下行标准差 """
        return {"下行标准差年化" : utils.get_value(self.decline_std()) * math.sqrt(52)}
    
    def sortino_ratio(self, risk_free_rate: float = 0.015) -> dict:
        """ 获取索提诺比率，无风险利率默认是 1.5% """
        return {"Sortino比率" : (utils.get_value(self.annual_return()) - risk_free_rate) / 
                                utils.get_value(self.decline_std_annualize())}
    
    def calmar_ratio(self) -> dict:
        """ 获取卡玛比 """
        return {"Calmar比率" : utils.get_value(self.annual_return()) / -utils.get_value(self.max_drawdown())}
    
    def is_new_high(self) -> dict:
        """ 最新日期是否创下新高？ """
        return {"是否创新高" : "是" if self.net_val[self.get_last_date()] >= self.net_val.max() else "否"}
    
    def days_until_new_high(self) -> dict:
        """ 未创新高的天数 """
        indicator_name = "未创新高的天数"
        return {indicator_name : (self.get_last_date() - self.net_val.idxmax()).days}
    
    def summary_indicators(self) -> dict: 
        """ 汇总除了 年度收益、月度收益及近期收益 之外的所有指标 """
        return {**self.cumulative_return(), **self.annual_return(), **self.max_drawdown(),
                **self.annual_volatility(), **self.sharpe_ratio(), **self.weekly_win_rate(),
                **self.this_week_return(), **self.max_drawdown_of_recent_year(), **self.max_weekly_drawdown(),
                **self.decline_std(), **self.decline_std_annualize(), **self.sortino_ratio(),
                **self.calmar_ratio(), **self.is_new_high(), **self.days_until_new_high()}
    
    def get_single_quantile(self, quantile: float, period_name: str) -> float:
        """
        获取指定滚动期限的收益率

        Args:
            quantile (float): 分位数，比如 0.25 代表求解0.25分位数，0.0代表求解最小值
            period_name (str): 滚动期限，必须是字符串 [半年, 一年, 二年, 三年, 五年] 之一
        """
        if period_name not in self.rolling_return_data.columns:
            raise ValueError(period_name, "必须是下列值之一:", self.rolling_return_data.columns)
        return self.rolling_return_data[period_name].quantile(quantile)
    
    def get_rolling_quantile_dataframe(self) -> pd.DataFrame:
        """ 获得滚动收益分位数表，即 最小/25分位/中位数/75分位数/最大 """
        quantile_list = [0.0, 0.25, 0.50, 0.75, 1.00]
        rolling_periods = self.rolling_return_data.columns
        result = pd.DataFrame()
        for period_name in rolling_periods:
            result[period_name] = [self.get_single_quantile(quantile, period_name) 
                                   for quantile in quantile_list]
        result.index = ["最小值", "25分位", "中位数", "75分位", "最大值"]
        result.index.name = "滚动收益"
        return result
    
    def get_earning_probability(self) -> pd.DataFrame:
        """ 获得盈利概率表 """
        prob_list = [0, 0.03, 0.05, 0.10, 0.12, 0.15, 0.18, 0.20]
        rolling_periods = self.rolling_return_data.columns
        result = pd.DataFrame()
        for period_name in rolling_periods:
            this_period_rolling: pd.Series = self.rolling_return_data[period_name]
            valid_count = this_period_rolling.count()
            result[period_name] = [this_period_rolling[this_period_rolling >= prob].count() / valid_count
                                   if valid_count != 0 else np.nan for prob in prob_list]
        result.index = ["{:.0%}".format(prob) for prob in prob_list]
        result.index.name = "盈利概率"
        return result
    
    def history_return_table(self, start_year: int = None) -> np.ndarray:
        """
        生成历史月度收益率表格每个单元格需要填充的内容

        Args:
            start_year (int, optional): 选择起始年份，如果没有选择，则从第一个净值日期所在年份开始计算. Defaults to None.
        
        Returns:
            np.ndarry: 返回 table 对应的矩阵，矩阵的行数和列数与 table 的行数和列数一致
        """
        self.err_check_start_year_ge_first_netval(start_year)
        # 范围是 [start_year, end_year] 两侧都是闭区间
        start_year = self.get_first_netval_date().year if start_year is None else start_year
        end_year = self.get_last_date().year
        return_matrix = []
        table_headers =  ["年份"] + [str(idx) + "月" for idx in range(1, 13)] + ["全年"]
        return_matrix.append(table_headers)
        for year in range(end_year, start_year - 1, -1):
            return_matrix.append(self.get_month_return_line(year))
        return np.array(return_matrix)
    
    def return_risk_table(self, start_year: int = None) -> np.ndarray:
        """
        生成表格：“收益风险指标”每个单元格需要填充的内容。
        注意：年度收益会包括基金净值数据第一年和最新一年的数据，即便这两个年份的收益率或许并不是全年的收益率。

        Args:
            start_year (int, optional): 选择起始年份，如果没有选择，则从第一个净值日期所在年份开始计算. Defaults to None.
        Returns:
            np.ndarray: 返回 table 对应的矩阵，矩阵的行数和列数与 table 的行数和列数一致
        """     
        self.err_check_start_year_ge_first_netval(start_year)

        table_contents = []
        table_headers = self.get_risk_table_headers()
        indicators_line = self.get_risk_table_header_indicators()
        table_contents.append(table_headers)
        table_contents.append(indicators_line)

        # 范围是 [start_year, end_year] 两侧都是闭区间
        start_year = self.get_first_netval_date().year if start_year is None else start_year
        end_year = self.get_last_date().year   
        table_contents += self.get_year_return_lines(start_year, end_year, len(table_headers) - 1)
        return np.array(table_contents)
    
    def get_risk_table_headers(self) -> list:
        """ 生成表格：“收益风险指标” 的表头 """
        return ["指标", "累计收益率", "年化收益率", "最大回撤", "年化波动率", "夏普比率", "周胜率"]
    
    def get_risk_table_header_indicators(self) -> list:
        """ 生成表格： “收益风险指标” 表头对应的数值"""
        indicators = {**self.cumulative_return(), **self.annual_return(), **self.max_drawdown(),
                **self.annual_volatility(), **self.sharpe_ratio(), **self.weekly_win_rate()}
        return [self.fund_name] +  [utils.round_decimal(indicators[key]) if "夏普" in key else 
                                               utils.decimal_to_pct(indicators[key]) for key in indicators.keys()]

    
    def get_year_return_lines(self, start_year: int, end_year: int, one_row_nums: int = 6) -> list:
        """
        获取 “收益风险指标” 表格中 年度/收益 年度/收益这些行，范围是从 [start_year, end_year] 

        Args:
            - start_year (int): 开始年份
            - end_year (int): 结束年份
            - one_row_nums (int, optional): 一行包含几年的收益率指标？ Defaults to 6.

        Returns:
            list:  “收益风险指标” 表格中 年度/收益 这些行需要填充的内容。注意：某些行没有填满，必须以空字符串""代替
        """
        year_nums = (end_year - start_year + 1)
        total_year_list = [year for year in range(end_year, start_year - 1, -1)]
        iter_nums = year_nums // one_row_nums + (0 if year_nums % one_row_nums == 0 else 1)
        result = []
        for iterator in range(iter_nums):
            year_list = total_year_list[iterator * one_row_nums : (iterator + 1) * one_row_nums]
            year_list += [None] * (one_row_nums - len(year_list))
            header_line = ["年度"] + ["" if elem is None else str(elem) + "年" for elem in year_list]
            value_line = [self.one_year_return(year) for year in year_list]
            value_header = "超额收益" if "超额" in self.fund_name else "收益"
            value_line = [value_header] + [utils.decimal_to_pct(elem) if not pd.isna(elem) else "" for elem in value_line]
            result.append(header_line)
            result.append(value_line)
        return result

    def get_month_return_line(self, year: int) -> list:
        """ 
        生成私募报告中“历史收益分析”某一行的数据
        获取某一年的月度收益率数据 + 全年数据(最后一年和净值日期开始的年份的“全年”列默认是空值) \n
        NOTE 基金净值首个月收益率和最新一个月的收益率可能并不是整个月的收益率，但是会被包括在下表中。
        比如：如果数据只到2023/10/21，那么2023年10月的收益率表示的是从10月初到2023/10/21日的收益率
        """
        monthly_return = [self.one_month_return(year, month) for month in range(1, 13)]
        monthly_return = ["-" if pd.isna(elem) else utils.decimal_to_pct(elem) for elem in monthly_return]
        # NOTE 这里简化处理：基金净值日期第一年和基金净值日期最后一年的“全年”列都是空值
        check_year: bool = (year == self.get_first_netval_date().year) or (year == self.get_last_date().year)
        yearly_return =  ["-"] if check_year else [utils.decimal_to_pct(self.one_year_return(year))]
        return [str(year)] + monthly_return + yearly_return
    
    def get_chart_data(self, index_data: pd.DataFrame) -> str:
        """
        获取绘制净值走势和回撤的数据[一般包含基金标准化净值，回撤，指数数据]，
        会从开始日期 或者 第一个净值日期 开始计算

        Args:
            - index_data(pd.DataFrame): 指数数据，可以是指数的原始数据，支持同时传入多个指数
        Returns:
            返回导出作图文件的名称
        """
        index_handler = ih.IndexHandler(index_data, self.get_first_netval_date())
        drawdown_col_name = self.get_column_name("回撤")
        standalized_col_name = self.get_column_name("标准化")
        the_fund_data = self.basic_data[[drawdown_col_name, 
                                         standalized_col_name]][self.basic_data.index >= self.get_first_netval_date()]
        the_fund_data.rename(columns = {standalized_col_name : utils.drop_suffix(standalized_col_name),
                                        drawdown_col_name : "最大回撤(右轴)"}, inplace = True)
        return the_fund_data.merge(index_handler.index_data, how = "left", left_index = True, right_index = True)
    
    def export_chart_data(self, merged_data: pd.Series) -> str:
        """ 将 get_chart_data 的返回结果进行导出，返回导出的作图文件的名称 """
        file_name = utils.generate_filename("__" + self.fund_name + "作图数据") 
        merged_data.to_excel("output/" + file_name)
        return file_name

    def get_analyze_text(self, start_year: int = None):
        """ 获取私募报告中要填写的分析文本

        模板：中金量化-贝叶斯稳健 1 号 2020 年 12 月以来累计收益 53.1%，年化收益 15.6%，产品正常运作以来最大回撤-6.7%，夏普 1.5，胜率 55.0%。
              分年度看，2021 年收益 37.2%、2022 年收益 4.5%、2023 年截至 12 月 1 日收益 6.8%。 
        
        Args:
            start_year(int): 可以指定在分析年度收益时，从哪年开始分析，可选参数。
        """
        self.err_check_start_year_ge_first_netval(start_year)

        blank_fill = "超额" if "超额" in self.fund_name else ""
        fund_name = self.fund_name.rstrip("-超额")
        start_time = self.get_first_netval_date().strftime("%Y年%m月")
        cumulative_return = utils.decimal_to_pct(list(self.cumulative_return().values())[0])
        annual_return = utils.decimal_to_pct(list(self.annual_return().values())[0])
        max_drawdown = utils.decimal_to_pct(list(self.max_drawdown().values())[0])
        sharpe_ratio = utils.round_decimal(list(self.sharpe_ratio().values())[0])
        weekly_win_rate = utils.decimal_to_pct(list(self.weekly_win_rate().values())[0])
        template_str = f"{fund_name}{start_time}以来累计{blank_fill}收益{cumulative_return}，" + \
                       f"年化{blank_fill}收益{annual_return}，产品正常运作以来{blank_fill}最大回撤{max_drawdown}，" + \
                       f"{blank_fill}夏普比率{sharpe_ratio}，{blank_fill}周胜率{weekly_win_rate}。"
        
        start_year = self.get_first_netval_date().year if start_year is None else start_year
        end_year = self.get_last_date().year
        start_day = self.get_first_netval_date().strftime("%m月%d日")
        start_year_return = utils.decimal_to_pct(self.one_year_return(start_year))

        if start_year == end_year: # NOTE 特殊情况，比如只有2023年需要分析的情况
            end_day = self.get_last_date().strftime("%m月%d日")
            return template_str + f"分年度看，{start_year}年从{start_day}截至{end_day}{blank_fill}收益{start_year_return}。"
        
        yearly_return_str = f"分年度看，{start_year}年从{start_day}到年底{blank_fill}收益{start_year_return}、" \
                            if start_year == self.get_first_netval_date().year else f"分年度看，{start_year}年{blank_fill}收益{start_year_return}、"
        for year in range(start_year + 1, end_year):
            year_return = utils.decimal_to_pct(self.one_year_return(year))
            yearly_return_str += f"{year}年{blank_fill}收益{year_return}、"
        end_day = self.get_last_date().strftime("%m月%d日")
        end_year_return = utils.decimal_to_pct(self.one_year_return(end_year))
        yearly_return_str += f"{end_year}年截至{end_day}{blank_fill}收益{end_year_return}。"
        return template_str + yearly_return_str
    
    def get_footnote_text(self, corp_name: str) -> str:
        """
        生成文本[数据来源：中金量化；指标计算区间为 2020 年 12 月 25 日至 2023 年 12 月 1 日。] \n
        NOTE BUG 注意：暂时不支持根据表格table的start_year动态调整指标计算区间的起始值

        Args:
            corp_name(str): 私募报告所涉及私募基金公司的名称
        """
        start_date_str = self.get_first_netval_date().strftime("%Y年%m月%d日")
        end_date_str = self.get_last_date().strftime("%Y年%m月%d日")
        return f"数据来源：{corp_name}，指标计算区间为{start_date_str}至{end_date_str}。"
        
        

