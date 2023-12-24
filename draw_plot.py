import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from datetime import date


def get_closest_val(val: float):
    """ 净值数据中，获取最接近某个 0.2 的下界/上界 """
    if val == 1:
        return 1
    if val < 1:
        result = 1
        while result > val + 1e-6:
            result -= 0.2
    if val > 1:
        result = 1
        while result < val - 1e-6:
            result += 0.2
    return round(result, 1)

def get_closest_percent(val: float, acc = 0.05):
    """ 回撤数据中，获取最接近某个 0.5% 的下界/上界 """
    if val == 0:
        return 0
    if val < 0:
        result = 0
        while result > val + 1e-7:
            result -= acc
    return round(result, 3)

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
    if _str.endswith(suffix):
        return _str[:-len(suffix)]
    if _str.endswith("-" + suffix):
        return _str[:-(len(suffix) + 1)]
    return _str

def get_best_interval(worst_drawdown: float, tick_num: int) -> float:
    """
    最大回撤的坐标刻度究竟该间隔 1% 2% 3% 4% 5% ...？
    该函数就是用来解决此问题的 

    Args:
        - worst_drawdown (float): 最大回撤，即回撤最小值
        - tick_num (int): 有多少个刻度线？

    Returns:
        float: 最大回撤右轴的最佳刻度间隔(负值)
    """
    for interval in range(1, 11):
        if - (tick_num - 1) * interval <= worst_drawdown * 100:
            return - interval / 100
    return -0.2

# 此字典储存了一些常用配色的 RGB值
the_color_map = {
    "deep_red" : np.array([192, 0, 0]) / 255,
    "orange" : np.array([255, 192, 0]) / 255,
    "deep_grey" : np.array([127]*3) / 255,
    "shallow_grey" : np.array([217] * 3) / 255
}

class GraphDrawer:
    def __init__(self, netval_data: pd.DataFrame, drawdown_data: pd.Series, fund_name: str):
        """
        此类之作用在于绘制出私募报告的标准图像[多条曲线是净值数据，阴影图表示回撤数据]

        Args:
            - netval_data (pd.DataFrame): 净值数据，包括基金净值和相关指数的净值数据，将会以左轴为纵轴
            - drawdown_data (pd.Series): 回撤数据，将会以右轴为纵轴
        """
        # 如果日期有任意一天不相等，就报错
        if not netval_data.index.equals(drawdown_data.index):
            raise ValueError("净值数据对应的日期序列须与回撤数据对应的日期序列完全一致")
        if type(netval_data) != pd.DataFrame or type(drawdown_data) != pd.Series:
            raise ValueError("输入数据类型错误，请看本函数注释")
        plt.rcParams['font.sans-serif']=['Kaiti'] # 用来正常显示楷体图例
        plt.rcParams['axes.unicode_minus']=False
        
        # 需要截断日期，以有基金净值的第一天开始画图
        self.netval_data = netval_data[netval_data.index >= netval_data.first_valid_index()]
        self.drawdown_data = drawdown_data[netval_data.index >= netval_data.first_valid_index()]
        self.fund_name = fund_name # 设置基金名称
        plt.figure(figsize = (10, 4))
        self.ax = plt.gca() # 获得绘图区域(左轴)
        self.ax2 = self.ax.twinx() # 获得绘图区域(右轴)
        self.xtick_nums = 12 # 横轴你想设置显示多少个日期？
        self.axis_font = "Arial" # 设置坐标轴字体
        self.axis_font_size = 10 # 设置坐标轴字体大小
        self.color_map = ["deep_red", "orange", "deep_grey"]
        self.x_values = np.arange(0, len(self.netval_data.index)) # 初始化横轴的所有坐标值
    
    def basic_set(self):
        """ 对绘图区的绘图次序、背景颜色进行初始设置 """
        self.ax.zorder = 1 # 设置左轴和右轴绘图的先后次序，右轴 ax2 先绘图，左轴 ax 后绘图
        self.ax2.zorder = 0 
        self.ax.set_facecolor("none") # 设置背景颜色是空
        self.ax2.set_facecolor("none")

    def set_xaxis_ticks(self) -> list:
        """ 设置横轴的tick，就是刻度线的所有位置 """
        x_ticks = list(map(lambda elem : int(elem), np.linspace(0, len(self.netval_data.index) - 1, self.xtick_nums)))
        self.ax.set_xticks(x_ticks)
        self.ax.set_xlim([min(self.x_values), max(self.x_values)])
        return x_ticks
    
    def set_left_axis_limit(self):
        """ 设置左轴数据范围[上下限]
        - 下限为 小于等于净值数据最小值 且 最接近0.2整数倍 的数据，例如 min 值为 0.71 对应的坐标轴下限为 0.6
        - 上限为 大于等于净值数据最大值 且 最接近0.2整数倍 的数据，例如 max 值为 1.66 对应的坐标轴上限为 1.8
        """
        left_lower = self.netval_data.min().min()
        left_upper = self.netval_data.max().max()
        self.ax.set_ylim([get_closest_val(left_lower), get_closest_val(left_upper)])
        self.ax.set_yticklabels(list(map(lambda elem : round(elem, 1), self.ax.get_yticks())), 
                                size = self.axis_font_size, fontproperties = self.axis_font)
        
    def set_right_axis_limit(self):
        """ 设置右轴数据范围[上下限]
        - 下限为 小于等于回撤数据最小值 且 最接近0.005(0.5%)整数倍 的数据，例如 min 值为 -2.68% 对应的坐标轴下限为 -3.0%
        - 上限为 显然，回撤数据的上限是0
        """
        right_upper = 0
        right_lower = self.drawdown_data.min()
        self.ax2.set_ylim([get_closest_percent(right_lower), right_upper])
        step = get_best_interval(self.ax2.get_ylim()[0], len(self.ax.get_yticks()))
        sub_y_ticks = np.arange(right_upper, right_upper + step * len(self.ax.get_yticks()), step)
        self.ax2.set_yticks(sub_y_ticks)
        self.ax2.set_yticklabels(list(map(lambda elem : "{:.0%}".format(elem), self.ax2.get_yticks())), 
                                 size = self.axis_font_size, fontproperties = self.axis_font)

    def set_left_ax(self):
        """ 设置左轴绘图对象 """
        self.set_left_axis_limit()
        self.ax.spines['top'].set_visible(False) #去掉上边框
        self.ax.spines['right'].set_visible(False) #去掉右边框
        self.ax.spines['left'].set_visible(False) #去掉左边框
        self.ax.spines['bottom'].set_position(('data',self.ax.get_ylim()[0]))

        self.ax.tick_params('x', direction = "in")
        self.ax.tick_params('y',left = False, right = False) # 去掉左轴刻度线
    
    def set_right_ax(self):
        """ 设置右轴绘图对象 """
        self.set_right_axis_limit()
        self.ax2.spines['top'].set_visible(False) #去掉上边框
        self.ax2.spines['top'].set_position(('data', self.ax2.get_ylim()[1])) # 设置右轴底部位置
        self.ax2.spines['right'].set_visible(False) #去掉右边框
        self.ax2.spines['left'].set_visible(False) #去掉左边框
        self.ax2.spines['bottom'].set_visible(False) # 去掉下边框

        self.ax2.tick_params('y', left = False, right = False) # 去掉左轴刻度线
        self.ax2.spines['bottom'].set_position(('data', self.ax2.get_ylim()[0])) 
    
    def plot_data(self):
        """ 将数据绘制到图像上 """
        for idx in range(len(self.netval_data.columns)):
            column_name = self.netval_data.columns[idx]
            self.ax.plot(self.x_values, self.netval_data[column_name].values, 
                         color = the_color_map[self.color_map[idx]], label = drop_suffix(column_name))    
        self.ax2.plot(self.x_values, self.drawdown_data.values, linewidth = 0.5, color = the_color_map["shallow_grey"]) 
        self.ax2.plot(self.x_values, [0] * len(self.drawdown_data.values), color = the_color_map["shallow_grey"], linewidth = 0.5) 
        self.ax2.fill_between(self.x_values, self.drawdown_data.values, [0] * len(self.drawdown_data), 
                              color = the_color_map["shallow_grey"], zorder = 2, label = "最大回撤（右轴）")
    
    def set_legend(self):
        """ 设置图例项 """
        #第一个参数是设置图例的具体坐标，第二个表示没有图例的边框
        self.ax.legend(loc = (0.3, 1.025),frameon=False,fontsize = 10, ncol = 3) 
        self.ax2.legend(loc = (0.05,1.025),frameon=False,fontsize = 10)

    def do_drawing(self):
        """ 调用上面的成员函数进行绘图 """
        self.basic_set()
        x_ticks = self.set_xaxis_ticks()
        self.set_left_ax()
        self.set_right_ax()
        self.plot_data()
        self.ax.set_xticklabels(list(map(lambda elem :  self.netval_data.index[elem], x_ticks)), 
            fontproperties = self.axis_font, size = 10, rotation = 45) # 设置横轴刻度
        self.ax.tick_params('x', direction = "in") # 横轴刻度线向内
        self.ax2.grid(axis='y', linestyle = "--",which = "major") # 设置横向虚线
        self.set_legend()
        output_path = "image/" + self.fund_name + "净值与回撤走势" + datetime.now().strftime("%y%m%d%H%M%S") + ".svg"
        plt.savefig(output_path, bbox_inches = 'tight')