import numpy as np
import pandas as pd
import time
import os

def get_time_str() -> str:
    """ 获得时间的字符串拼接，作为文件名后缀，这是为了防止输出文件被频繁覆盖 """
    result: str = ""
    result += "{:04d}".format(time.gmtime()[0])
    for elem in time.gmtime()[1:6]:
        result += "{:02d}".format(elem)
    return result

def get_file_path(file_name: str = "计算结果", output_dir: str = "output") -> str:
    """ 获取文件输出路径，文件名是 "计算结果 + 楼上的函数的时间字符串 + .xlsx" """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    file_name += ("-" + get_time_str() + ".xlsx")
    return output_dir + "/" + file_name

def data_clear(data: pd.DataFrame,name_fund,time_lis):  #data为array形式
    data_change= data.apply(pd.to_numeric, errors = 'coerce')     # 能转数值就转数值，不能的话就转 Nan
    data_change=data_change.interpolate(method='linear', axis=0) # 线性插值填充
    data_array=np.array(data) # 转为了 array类 型
    data_change_array=np.array(data_change)
    storage_lis=[] # 记录data_change 与 data_array 不一致的产品名称，对应日期，原始data数据，改变后的数据
    # NOTE storage_lis 这个变量似乎没有在后文的计算函数里面使用过。
    for i in range(data.shape[0]): # BUG 可能出现问题的地方：这里会把二者皆为 np.nan 的数据也判定为不相等，加入 storage_lis
        for j in range(data.shape[1]):
            if data_array[i][j]!=data_change_array[i][j]:
                storage_lis.append([time_lis[i],name_fund[j],data_array[i][j],data_change_array[i][j]])
    return data_change,data_change_array,storage_lis

def drawdown(data):
    data=data.values
    max_drawdown = []
    peak = None  # 初始化峰值为None
    for val in data:
        if pd.isna(val):  # 检查是否为有效数据
            max_drawdown.append(np.nan)
        else:
            if pd.isna(peak):
                peak = val
            elif val > peak:
                peak = val
            dd = (val - peak) / peak
            max_drawdown.append(dd)
    return pd.DataFrame(max_drawdown)

def match_recent_day(data,days):
    return ((data.iloc[-1]-data).apply(lambda x: x.days)-days).abs().idxmin()

def convert_to_date(elem):
    """ 将 datetime 转化为 date """
    try:
        return elem.date()
    except:
        return elem

def index_caculate(data_change,name_fund,time_lis,time_build):
    result_spf=pd.DataFrame()
    result_spf['时间']=time_lis
    for i in range(len(name_fund)):
        result_spf[name_fund[i]]=data_change[name_fund[i]]
        result_spf[name_fund[i]+'周度收益率']= result_spf[name_fund[i]].pct_change()  # 计算月度收益率
        result_spf[name_fund[i] + '最大回撤']=drawdown(result_spf[name_fund[i]])
        result_spf[name_fund[i] + '标准化'] = data_change[name_fund[i]]/(data_change[name_fund[i]][data_change[name_fund[i]].first_valid_index()])
    returns_spf=pd.DataFrame()
    for i in range(len(name_fund)):
        returns_spf[name_fund[i]+'周度收益率']=result_spf[name_fund[i]+'周度收益率']
    drawdown_spf = pd.DataFrame()
    for i in range(len(name_fund)):
        drawdown_spf[name_fund[i] + '最大回撤'] = result_spf[name_fund[i] + '最大回撤']

    #添加一行空值以及数据分析指标一行
    result_spf.loc[result_spf.shape[0]]=[np.nan]*result_spf.shape[1]
    result_spf.loc[result_spf.shape[0]] =['数据分析指标']+ [np.nan] * (result_spf.shape[1]-1)

    # 年度收益率
    for year in pd.Series(time_lis).dt.year.unique():
        temp_lis=[]
        for i in range(len(name_fund)):
            returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
            mask = (pd.Series(time_lis).dt.year == year)
            annual_return = (1 + returns[mask]).prod() - 1 # 寻找的是所有时间中以2017开头的那种
            temp_lis=temp_lis+[annual_return,np.nan,np.nan,np.nan]
        result_spf.loc[result_spf.shape[0]] = [str(year)+'年'] + temp_lis

    #今年以来,成立以来，本周收益率
    temp_lis_this_year=[]
    temp_lis_start_end= []
    temp_lis_weekly_return = []

    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        mask = (pd.Series(time_lis).dt.year == pd.Series(time_lis).dt.year.unique()[-1])
        this_year=(1 + returns[mask]).prod() - 1
        temp_lis_this_year = temp_lis_this_year + [this_year, np.nan, np.nan,np.nan]

        start_end_year = (1 + returns).prod() - 1
        temp_lis_start_end= temp_lis_start_end + [start_end_year, np.nan, np.nan,np.nan]

        temp_lis_weekly_return = temp_lis_weekly_return + [returns_spf[name_fund[i] + '周度收益率'].iloc[-1], np.nan, np.nan,np.nan]
    result_spf.loc[result_spf.shape[0]] = ['今年以来'] + temp_lis_this_year
    result_spf.loc[result_spf.shape[0]] = ['成立以来'] + temp_lis_start_end
    result_spf.loc[result_spf.shape[0]] = ['本周'] + temp_lis_weekly_return


    #是否创新高
    temp_lis = []
    for i in range(len(name_fund)):
        sign=data_change[name_fund[i]].iloc[-1] # 最后一个净值数据
        max_sign=data_change[name_fund[i]].max()
        if sign==max_sign:
            temp_lis=temp_lis + ['是', np.nan, np.nan,np.nan]
        else:
            temp_lis = temp_lis + ['否', np.nan, np.nan,np.nan]
    result_spf.loc[result_spf.shape[0]] = ['是否创新高'] + temp_lis

    #未创新高的天
    temp_lis = []
    for i in range(len(name_fund)):
        # BUG 这里不对，因为时间序列是以周为单位的，所以这里得出的是93周未创新高，而不是679天未创新高
        sign = (time_lis[-1] - time_lis[data_change[name_fund[i]].idxmax()]).days
        temp_lis = temp_lis + [sign, np.nan, np.nan,np.nan]

    result_spf.loc[result_spf.shape[0]] = ['未创新高的天'] + temp_lis

    # 近一月,三月，一年，两年，三年收益率
    temp_lis_one_month = []
    temp_lis_three_month = []
    temp_lis_six_month = []
    temp_lis_one_year = []
    temp_lis_two_year = []
    temp_lis_three_year = []
    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        # print(returns)
        # print(returns[(match_recent_day(pd.Series(time_lis),30)+1):])
        one_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),30)+1):]).prod() - 1 if len(returns) >=4  else np.nan
        three_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),90)+1):]).prod() - 1 if len(returns) >= 12 else np.nan
        six_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),365//2)+1):]).prod() - 1 if len(returns) >= 26 else np.nan        
        one_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),365)+1):]).prod() - 1 if len(returns) >= 52 else np.nan
        two_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),730)+1):]).prod() - 1 if len(returns) >= 104 else np.nan
        three_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis),1095)+1):]).prod() - 1 if len(returns) >=156 else np.nan

        temp_lis_one_month = temp_lis_one_month+[one_month , np.nan, np.nan,np.nan]
        temp_lis_three_month = temp_lis_three_month+[three_month, np.nan, np.nan,np.nan]
        temp_lis_six_month = temp_lis_six_month+[six_month, np.nan, np.nan,np.nan]
        temp_lis_one_year = temp_lis_one_year+[one_year, np.nan, np.nan,np.nan]
        temp_lis_two_year = temp_lis_two_year+[two_year, np.nan, np.nan,np.nan]
        temp_lis_three_year = temp_lis_three_year+[three_year, np.nan, np.nan,np.nan]
    result_spf.loc[result_spf.shape[0]] = ['近一个月'] + temp_lis_one_month
    result_spf.loc[result_spf.shape[0]] = ['近三个月'] +temp_lis_three_month
    result_spf.loc[result_spf.shape[0]] = ['近六个月'] +temp_lis_six_month
    result_spf.loc[result_spf.shape[0]] = ['近一年'] + temp_lis_one_year
    result_spf.loc[result_spf.shape[0]] = ['近两年'] + temp_lis_two_year
    result_spf.loc[result_spf.shape[0]] = ['近三年'] + temp_lis_three_year


    #年化收益率，收益率标准差，年化波动率,历史最大回撤，过去一年最大回撤，最大周度回撤，下行标准差，下行标准差年化，周胜率
    #夏普比率，sortino比率，calmar比率
    temp_lis_1 = []
    temp_lis_2 = []
    temp_lis_3 = []
    temp_lis_history_drawdown=[]
    temp_lis_year_drawdown = []
    temp_lis_max_week_drawdown = []
    temp_lis_down_std = []
    temp_lis_down_annual_std = []
    temp_lis_week_win = []
    temp_lis_sharpe = []
    temp_lis_sortino_ratio=[]
    temp_lis_calmar_ratio = []


    risk_free_rate = 0.015
    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        temp_returns=pow((1 + returns).prod(),365 / (time_lis[-1] - time_build[i]).days)-1 # NOTE 与 EXCEL 中的计算逻辑有出入，但也是合理的
        temp_lis_1 = temp_lis_1 + [temp_returns, np.nan, np.nan,np.nan]

        temp_std=returns.std()
        temp_lis_2 = temp_lis_2 + [temp_std, np.nan, np.nan,np.nan]

        temp_year_std=temp_std * pow(52, 0.5)
        temp_lis_3 = temp_lis_3 + [temp_year_std, np.nan, np.nan,np.nan]

        drawdowns = drawdown_spf[name_fund[i] + '最大回撤'].dropna().min()
        temp_lis_history_drawdown = temp_lis_history_drawdown + [drawdowns, np.nan, np.nan,np.nan]

        drawdowns_1 = drawdown_spf[name_fund[i] + '最大回撤'][-51:].dropna().min() # NOTE 这是在算过去一年最大回撤吗？应该是52周而不是12周
        temp_lis_year_drawdown = temp_lis_year_drawdown + [drawdowns_1, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna().min() # 最大周度回撤
        temp_lis_max_week_drawdown = temp_lis_max_week_drawdown + [returns, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        # NOTE 周报里面下行标准差的计算公式也是这么写的
        downside_std = (sum(returns[returns < 0] ** 2) / (len(returns) - 1)) ** 0.5 # NOTE 标准差公式建议直接使用 returns[returns < 0].std()
        annualized_downside_std = downside_std * np.sqrt(52)
        temp_lis_down_std = temp_lis_down_std + [downside_std, np.nan, np.nan,np.nan]
        temp_lis_down_annual_std = temp_lis_down_annual_std+ [annualized_downside_std, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        weekly_win_rate = (returns > 0).mean()
        temp_lis_week_win = temp_lis_week_win + [weekly_win_rate, np.nan, np.nan,np.nan]

        sharpe_ratio = (temp_returns - risk_free_rate) / temp_year_std
        temp_lis_sharpe = temp_lis_sharpe + [sharpe_ratio, np.nan, np.nan,np.nan]

        sortino_ratio = (temp_returns - risk_free_rate) / annualized_downside_std
        temp_lis_sortino_ratio= temp_lis_sortino_ratio + [sortino_ratio, np.nan, np.nan,np.nan]

        calmar_ratio = - temp_returns / drawdowns
        temp_lis_calmar_ratio = temp_lis_calmar_ratio + [calmar_ratio, np.nan, np.nan,np.nan]

    result_spf.loc[result_spf.shape[0]] = ['年化收益率'] + temp_lis_1
    result_spf.loc[result_spf.shape[0]] = ['收益率标准差'] + temp_lis_2
    result_spf.loc[result_spf.shape[0]] = ['年化波动率'] + temp_lis_3
    result_spf.loc[result_spf.shape[0]] = ['历史最大回撤'] + temp_lis_history_drawdown
    result_spf.loc[result_spf.shape[0]] = ['过去一年最大回撤'] + temp_lis_year_drawdown
    result_spf.loc[result_spf.shape[0]] = ['最大周度回撤'] + temp_lis_max_week_drawdown
    result_spf.loc[result_spf.shape[0]] = ['下行标准差'] + temp_lis_down_std
    result_spf.loc[result_spf.shape[0]] = ['下行标准差年化'] + temp_lis_down_annual_std
    result_spf.loc[result_spf.shape[0]] = ['周胜率'] + temp_lis_week_win
    result_spf.loc[result_spf.shape[0]] = ['风险调整后收益比率：'] + len(temp_lis_week_win) * [np.nan]
    result_spf.loc[result_spf.shape[0]] = ['夏普比率'] + temp_lis_sharpe
    result_spf.loc[result_spf.shape[0]] = ['Sortino比率'] + temp_lis_sortino_ratio
    result_spf.loc[result_spf.shape[0]] = ['Calmar比率'] + temp_lis_calmar_ratio
    result_spf["时间"] = result_spf["时间"].apply(convert_to_date)
    result_spf.to_excel(get_file_path(file_name = "计算结果-非指增"), index=False)

def Excess_returns(fund_returns,benchmark_returns):
    return (fund_returns+1)/(benchmark_returns+1)

def Enhanced_index_caculate(data_change,name_fund,time_lis,time_build,index_lis):
    result_spf = pd.DataFrame()
    result_spf['时间'] = time_lis
    for i in range(len(name_fund)):
        result_spf[name_fund[i]] = data_change[name_fund[i]]
        result_spf[name_fund[i] + '周度收益率'] = result_spf[name_fund[i]].pct_change()  # 计算月度收益率
        result_spf[name_fund[i] + '最大回撤'] = drawdown(result_spf[name_fund[i]])
        result_spf[name_fund[i] + '标准化'] =data_change[name_fund[i]]/(data_change[name_fund[i]][data_change[name_fund[i]].first_valid_index()])


        temp_series=Excess_returns(result_spf[name_fund[i] + '周度收益率'],index_lis.pct_change())
        temp_sign=temp_series.first_valid_index()
        lis_temp=[np.nan for j in range(temp_sign-1)]+[1]
        for p in range(len(data_change)-temp_sign):

            lis_temp.append(lis_temp[-1]*(temp_series[p+temp_sign]))
        result_spf[name_fund[i] + '超额收益'] =lis_temp

        result_spf[name_fund[i] + '超额收益收益率']=result_spf[name_fund[i] + '超额收益'].pct_change()  # 计算月度收益率
        result_spf[name_fund[i] + '超额收益最大回撤']=drawdown(result_spf[name_fund[i] + '超额收益'])

    returns_spf = pd.DataFrame()
    for i in range(len(name_fund)):
        returns_spf[name_fund[i] + '周度收益率'] = result_spf[name_fund[i] + '周度收益率']
        returns_spf[name_fund[i] + '超额收益收益率'] = result_spf[name_fund[i] + '超额收益收益率']

    drawdown_spf = pd.DataFrame()
    for i in range(len(name_fund)):
        drawdown_spf[name_fund[i] + '最大回撤'] = result_spf[name_fund[i] + '最大回撤']
        drawdown_spf[name_fund[i] + '超额收益最大回撤'] = result_spf[name_fund[i] + '超额收益最大回撤']

    excess_return_spf=pd.DataFrame()
    for i in range(len(name_fund)):
        excess_return_spf[name_fund[i] + '超额收益'] = result_spf[name_fund[i] + '超额收益']




    # 添加一行空值以及数据分析指标一行
    result_spf.loc[result_spf.shape[0]] = [np.nan] * result_spf.shape[1]
    result_spf.loc[result_spf.shape[0]] = ['数据分析指标'] + [np.nan] * (result_spf.shape[1] - 1)

    # 年度收益率
    for year in pd.Series(time_lis).dt.year.unique():
        temp_lis = []
        for i in range(len(name_fund)):
            returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
            mask = (pd.Series(time_lis).dt.year == year)
            annual_return = (1 + returns[mask]).prod() - 1
        #超额收益部分
            returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()
            mask_1 = (pd.Series(time_lis).dt.year == year)
            annual_return_1 = (1 + returns_1[mask]).prod() - 1
            temp_lis = temp_lis + [annual_return, np.nan, np.nan, np.nan,annual_return_1,np.nan,np.nan]
        result_spf.loc[result_spf.shape[0]] = [str(year) + '年'] + temp_lis


    # 今年以来,成立以来，本周
    temp_lis_this_year = []
    temp_lis_start_end = []
    temp_lis_weekly_return = []

    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        mask = (pd.Series(time_lis).dt.year == pd.Series(time_lis).dt.year.unique()[-1])
        this_year = (1 + returns[mask]).prod() - 1
        temp_lis_this_year = temp_lis_this_year + [this_year, np.nan, np.nan,np.nan]

        start_end_year = (1 + returns).prod() - 1
        temp_lis_start_end = temp_lis_start_end + [start_end_year, np.nan, np.nan,np.nan]

        temp_lis_weekly_return = temp_lis_weekly_return + [returns_spf[name_fund[i] + '周度收益率'].iloc[-1], np.nan, np.nan,np.nan]

       #超额收益部分
        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()
        mask = (pd.Series(time_lis).dt.year == pd.Series(time_lis).dt.year.unique()[-1])
        this_year_1 = (1 + returns_1[mask]).prod() - 1
        temp_lis_this_year = temp_lis_this_year + [this_year_1, np.nan, np.nan]

        start_end_year_1 = (1 + returns_1).prod() - 1
        temp_lis_start_end = temp_lis_start_end + [start_end_year_1, np.nan, np.nan]

        temp_lis_weekly_return = temp_lis_weekly_return + [returns_spf[name_fund[i] + '超额收益收益率'].iloc[-1], np.nan, np.nan]
    result_spf.loc[result_spf.shape[0]] = ['今年以来'] + temp_lis_this_year
    result_spf.loc[result_spf.shape[0]] = ['成立以来'] + temp_lis_start_end
    result_spf.loc[result_spf.shape[0]] = ['本周'] + temp_lis_weekly_return


    # 是否创新高
    temp_lis = []
    for i in range(len(name_fund)):
        sign = data_change[name_fund[i]].iloc[-1]
        max_sign = data_change[name_fund[i]].max()
        if sign == max_sign:
            temp_lis = temp_lis + ['是', np.nan, np.nan,np.nan]
        else:
            temp_lis = temp_lis + ['否', np.nan, np.nan,np.nan]

        #超额收益部分
        sign_1 = excess_return_spf[name_fund[i]+'超额收益'].iloc[-1]
        max_sign_1 = excess_return_spf[name_fund[i]+'超额收益'].max()
        if sign == max_sign:
            temp_lis = temp_lis + ['是', np.nan, np.nan]
        else:
            temp_lis = temp_lis + ['否', np.nan, np.nan]

    result_spf.loc[result_spf.shape[0]] = ['是否创新高'] + temp_lis


    # 未创新高的天
    temp_lis = []
    for i in range(len(name_fund)):
        sign = (time_lis[-1] - time_lis[data_change[name_fund[i]].idxmax()]).days # BUG 未创新高的天数，这里算成了未创新高的周
        temp_lis = temp_lis + [sign, np.nan, np.nan,np.nan]

    #超额收益部分
        sign_1 = (time_lis[-1] - time_lis[excess_return_spf[name_fund[i] + '超额收益'].idxmax()]).days
        temp_lis = temp_lis + [sign_1, np.nan, np.nan] # BUG 未创新高的天数，这里算成了未创新高的周
    result_spf.loc[result_spf.shape[0]] = ['未创新高的天'] + temp_lis




    # 近一月,三月，一年，两年，三年收益率
    temp_lis_one_month = []
    temp_lis_three_month = []
    temp_lis_six_month = []
    temp_lis_one_year = []
    temp_lis_two_year = []
    temp_lis_three_year = []
    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        one_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 30) + 1):]).prod() - 1 if len(returns) >= 4 else np.nan
        three_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 90) + 1):]).prod() - 1 if len(returns) >= 12 else np.nan
        six_month = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 365//2) + 1):]).prod() - 1 if len(returns) >= 26 else np.nan
        one_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 365) + 1):]).prod() - 1 if len(returns) >= 52 else np.nan
        two_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 730) + 1):]).prod() - 1 if len(returns) >= 104 else np.nan
        three_year = (1 + returns[returns.index.get_loc(match_recent_day(pd.Series(time_lis), 1095) + 1):]).prod() - 1 if len(returns) >= 156 else np.nan

        temp_lis_one_month = temp_lis_one_month + [one_month, np.nan, np.nan,np.nan]
        temp_lis_three_month = temp_lis_three_month + [three_month, np.nan, np.nan,np.nan]
        temp_lis_six_month = temp_lis_six_month + [six_month, np.nan, np.nan,np.nan]
        temp_lis_one_year = temp_lis_one_year + [one_year, np.nan, np.nan,np.nan]
        temp_lis_two_year = temp_lis_two_year + [two_year, np.nan, np.nan,np.nan]
        temp_lis_three_year = temp_lis_three_year + [three_year, np.nan, np.nan,np.nan]

        #超额收益部分
        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()
        one_month_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 30) + 1):]).prod() - 1 if len(
            returns_1) >= 4 else np.nan
        three_month_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 90) + 1):]).prod() - 1 if len(
            returns_1) >= 12 else np.nan
        six_month_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 365 // 2) + 1):]).prod() - 1 if len(
            returns_1) >= 26 else np.nan
        one_year_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 365) + 1):]).prod() - 1 if len(
            returns_1) >= 52 else np.nan
        two_year_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 730) + 1):]).prod() - 1 if len(
            returns_1) >= 104 else np.nan
        three_year_1 = (1 + returns_1[returns_1.index.get_loc(match_recent_day(pd.Series(time_lis), 1095) + 1):]).prod() - 1 if len(
            returns_1) >= 156 else np.nan

        temp_lis_one_month = temp_lis_one_month + [one_month_1, np.nan, np.nan]
        temp_lis_three_month = temp_lis_three_month + [three_month_1, np.nan, np.nan]
        temp_lis_six_month = temp_lis_six_month + [six_month_1, np.nan, np.nan]
        temp_lis_one_year = temp_lis_one_year + [one_year_1, np.nan, np.nan]
        temp_lis_two_year = temp_lis_two_year + [two_year_1, np.nan, np.nan]
        temp_lis_three_year = temp_lis_three_year + [three_year_1, np.nan, np.nan]

    result_spf.loc[result_spf.shape[0]] = ['近一个月'] + temp_lis_one_month
    result_spf.loc[result_spf.shape[0]] = ['近三个月'] + temp_lis_three_month
    result_spf.loc[result_spf.shape[0]] = ['近六个月'] + temp_lis_six_month
    result_spf.loc[result_spf.shape[0]] = ['近一年'] + temp_lis_one_year
    result_spf.loc[result_spf.shape[0]] = ['近两年'] + temp_lis_two_year
    result_spf.loc[result_spf.shape[0]] = ['近三年'] + temp_lis_three_year


    # 年化收益率，收益率标准差，年化波动率,历史最大回撤，过去一年最大回撤，最大周度回撤，下行标准差，下行标准差年化，周胜率
    # 夏普比率，sortino比率，calmar比率
    temp_lis_1 = []
    temp_lis_2 = []
    temp_lis_3 = []
    temp_lis_history_drawdown = []
    temp_lis_year_drawdown = []
    temp_lis_max_week_drawdown = []
    temp_lis_down_std = []
    temp_lis_down_annual_std = []
    temp_lis_week_win = []
    temp_lis_sharpe = []
    temp_lis_sortino_ratio = []
    temp_lis_calmar_ratio = []

    risk_free_rate = 0.015
    for i in range(len(name_fund)):
        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        temp_returns = pow((1 + returns).prod(), 365 / (time_lis[-1] - time_build[i]).days) - 1
        temp_lis_1 = temp_lis_1 + [temp_returns, np.nan, np.nan,np.nan]

        temp_std = returns.std()
        temp_lis_2 = temp_lis_2 + [temp_std, np.nan, np.nan,np.nan]

        temp_year_std = temp_std * pow(52, 0.5)
        temp_lis_3 = temp_lis_3 + [temp_year_std, np.nan, np.nan,np.nan]

        drawdowns = drawdown_spf[name_fund[i] + '最大回撤'].dropna().min()
        temp_lis_history_drawdown = temp_lis_history_drawdown + [drawdowns, np.nan, np.nan,np.nan]
        # BUG 过去一年：应该取52周而不是12周
        drawdowns_1 = drawdown_spf[name_fund[i] + '最大回撤'][-51:].dropna().min() 
        temp_lis_year_drawdown = temp_lis_year_drawdown + [drawdowns_1, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna().min()
        temp_lis_max_week_drawdown = temp_lis_max_week_drawdown + [returns, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        downside_std = (sum(returns[returns < 0] ** 2) / (len(returns) - 1)) ** 0.5
        annualized_downside_std = downside_std * np.sqrt(52)
        temp_lis_down_std = temp_lis_down_std + [downside_std, np.nan, np.nan,np.nan]
        temp_lis_down_annual_std = temp_lis_down_annual_std + [annualized_downside_std, np.nan, np.nan,np.nan]

        returns = returns_spf[name_fund[i] + '周度收益率'].dropna()
        weekly_win_rate = (returns > 0).mean()
        temp_lis_week_win = temp_lis_week_win + [weekly_win_rate, np.nan, np.nan,np.nan]

        sharpe_ratio = (temp_returns - risk_free_rate) / temp_year_std
        temp_lis_sharpe = temp_lis_sharpe + [sharpe_ratio, np.nan, np.nan,np.nan]

        sortino_ratio = (temp_returns - risk_free_rate) / annualized_downside_std
        temp_lis_sortino_ratio = temp_lis_sortino_ratio + [sortino_ratio, np.nan, np.nan,np.nan]

        calmar_ratio = - temp_returns / drawdowns
        temp_lis_calmar_ratio = temp_lis_calmar_ratio + [calmar_ratio, np.nan, np.nan,np.nan]


        #超额收益部分
        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()

        enhance_begin_time = max(time_lis[index_lis.first_valid_index()], time_build[i])
        temp_returns_1 = pow((1 + returns_1).prod(), 365 / (time_lis[-1] - enhance_begin_time).days) - 1
        temp_lis_1 = temp_lis_1 + [temp_returns_1, np.nan, np.nan]

        temp_std_1 = returns_1.std()
        temp_lis_2 = temp_lis_2 + [temp_std_1, np.nan, np.nan]

        temp_year_std_1 = temp_std_1 * pow(52, 0.5)
        temp_lis_3 = temp_lis_3 + [temp_year_std_1, np.nan, np.nan]

        drawdowns_11 = drawdown_spf[name_fund[i] + '超额收益最大回撤'].dropna().min()
        temp_lis_history_drawdown = temp_lis_history_drawdown + [drawdowns_11, np.nan, np.nan]
        # NOTE 过去一年：应该取52周而不是12周
        drawdowns_111 = drawdown_spf[name_fund[i] + '超额收益最大回撤'][-51:].dropna().min() 
        temp_lis_year_drawdown = temp_lis_year_drawdown + [drawdowns_111, np.nan, np.nan] 

        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna().min()
        temp_lis_max_week_drawdown = temp_lis_max_week_drawdown + [returns_1, np.nan, np.nan]

        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()
        downside_std_1 = (sum(returns_1[returns_1 < 0] ** 2) / (len(returns_1) - 1)) ** 0.5
        annualized_downside_std_1 = downside_std_1 * np.sqrt(52)
        temp_lis_down_std = temp_lis_down_std + [downside_std_1, np.nan, np.nan]
        temp_lis_down_annual_std = temp_lis_down_annual_std + [annualized_downside_std_1, np.nan, np.nan]

        returns_1 = returns_spf[name_fund[i] + '超额收益收益率'].dropna()
        weekly_win_rate_1 = (returns_1 > 0).mean()
        temp_lis_week_win = temp_lis_week_win + [weekly_win_rate_1, np.nan, np.nan]

        sharpe_ratio_1 = (temp_returns_1 - risk_free_rate) / temp_year_std_1
        # print(sharpe_ratio_1,temp_returns_1,temp_year_std_1)
        temp_lis_sharpe = temp_lis_sharpe + [sharpe_ratio_1, np.nan, np.nan]

        sortino_ratio_1 = (temp_returns_1 - risk_free_rate) / annualized_downside_std_1
        temp_lis_sortino_ratio = temp_lis_sortino_ratio + [sortino_ratio_1, np.nan, np.nan]

        calmar_ratio_1 = - temp_returns_1 / drawdowns_11 
        temp_lis_calmar_ratio = temp_lis_calmar_ratio + [calmar_ratio_1, np.nan, np.nan]


    result_spf.loc[result_spf.shape[0]] = ['年化收益率'] + temp_lis_1
    result_spf.loc[result_spf.shape[0]] = ['收益率标准差'] + temp_lis_2
    result_spf.loc[result_spf.shape[0]] = ['年化波动率'] + temp_lis_3
    result_spf.loc[result_spf.shape[0]] = ['历史最大回撤'] + temp_lis_history_drawdown
    result_spf.loc[result_spf.shape[0]] = ['过去一年最大回撤'] + temp_lis_year_drawdown
    result_spf.loc[result_spf.shape[0]] = ['最大周度回撤'] + temp_lis_max_week_drawdown
    result_spf.loc[result_spf.shape[0]] = ['下行标准差'] + temp_lis_down_std
    result_spf.loc[result_spf.shape[0]] = ['下行标准差年化'] + temp_lis_down_annual_std
    result_spf.loc[result_spf.shape[0]] = ['周胜率'] + temp_lis_week_win
    result_spf.loc[result_spf.shape[0]] = ['风险调整后收益比率：'] + len(temp_lis_week_win) * [np.nan]
    result_spf.loc[result_spf.shape[0]] = ['夏普比率'] + temp_lis_sharpe
    result_spf.loc[result_spf.shape[0]] = ['Sortino比率'] + temp_lis_sortino_ratio
    result_spf.loc[result_spf.shape[0]] = ['Calmar比率'] + temp_lis_calmar_ratio
    result_spf["时间"] = result_spf["时间"].apply(convert_to_date)
    result_spf.to_excel(get_file_path(file_name = "计算结果-指增"), index=False)

def calculate_index():
    """ 调用本文件的各种函数，进行指标计算，并导出数据 """
    # NOTE 这里应该输入数据的相对路径或者绝对路径
    # NOTE 建议使用相对路径而不是绝对路径，把数据放到该项目的文件夹里，大家都轻松
    # 指数数据地址
    tt = "data/指数数据.xlsx"
    # 指增基金数据地址
    pp = "data/指增测试数据.xlsx"
    # 非指增基金数据地址
    qq = "data/非指增测试数据.xlsx"

    sign=input('如果您需要处理的数据为指增数据请输入1，反之输入2：')
    if sign==str(1):
        # 读取指数数据
        index_spf=pd.read_excel(tt, index_col=False, header=0).iloc[:,:]
        print(list(index_spf.columns[1:]))
        ss=input('请问您需要选择的指数为上面指数列表里面的哪一个指数，请输入其名称：')

        spf = pd.read_excel(pp, index_col=False, header=0).iloc[:, :]
        name_fund = list(spf.columns)[1:]
        time_build = pd.to_datetime(spf.iloc[0, 1:])
        time_lis = [i[0] for i in spf.iloc[1:, 0:1].values]
        data = spf.iloc[1:, 1:]
        data = data.reset_index(drop=True)
        data_change, data_change_array, storage_lis = data_clear(data, name_fund, time_lis)
        Enhanced_index_caculate(data_change,name_fund,time_lis,time_build,index_spf[ss])

    else:
        # NOTE 这样读入数据后，首列列名是 Unnamed
        spf = pd.read_excel(qq, index_col=False, header=0).iloc[:, :]
        # 三个产品的名称
        name_fund = list(spf.columns)[1:]
        # 将时间转换为日期
        time_build = pd.to_datetime(spf.iloc[0, 1:])
        # 一个 datetime.datetime 列表
        time_lis = [i[0] for i in spf.iloc[1:, 0:1].values]
        data = spf.iloc[1:, 1:]
        # 一个神奇操作：data中的数据是不包含日期数据的
        data = data.reset_index(drop=True)
        data_change, data_change_array, storage_lis = data_clear(data, name_fund, time_lis)
        index_caculate(data_change, name_fund, time_lis, time_build)

if __name__ == "__main__":
    calculate_index()


