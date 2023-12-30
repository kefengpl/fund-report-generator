""" 此文件用于处理 Excel 绘图 """
# 引入第三方包
import win32com.client as win32

# 引入自己写的模块
import utils

AREA_PLOT = 76 # 面积图代号：76
LINE_PLOT = 4 # 折线图代号：4

class ExcelChartHandler:
    def __init__(self, file_path: str, visible: bool = True, 
                 chart_height: float = 40 * 7.8, chart_width: float = 40 * 22.5):
        """
        此类用于处理 Excel 作图。要求最左侧一列是[日期]， 前面若干列是[净值]，
        最后一列必须是[回撤]，且列名包含[回撤]两个字。不得包括其它数据。

        Args:
            - file_path (str): _description_
            - visible (bool, optional): 你是否希望打开 Excel，观察此绘图过程？
            - chart_height(float): 图表高度
            - chart_width(float): 图表宽度
        """
        self.excel_app = win32.Dispatch('Excel.Application') # 创建 Excel APP 应用程序
        self.excel_app.Visible = visible #显示 Excel 程序
        self.excel_file = self.excel_app.Workbooks.Open(file_path) # 打开我的Excel
        self.excel_sheet = self.excel_file.Worksheets("Sheet1") # 获取 Sheet1 对象
        self.used_range = self.excel_sheet.UsedRange # 获得使用区域对象
        self.n_rows = self.used_range.Rows.Count # 使用区域一共有多少行？
        self.n_cols = self.used_range.Columns.Count # 使用区域一共多少列？
        self.line_color_list: list =  list(utils.color_dict.keys())[:self.n_cols - 1] # 绘制线条的颜色列表，以 utils.color_dict 里面的颜色从头开始取
        self.excel_sheet.Shapes.AddChart(Width = chart_width,Height = chart_height).Select() # 对该sheet对象的所有列添加图表
        self.chart = self.excel_file.ActiveChart # 获取当前选中的图表
        self.chart.SetSourceData(self.excel_sheet.UsedRange) # 设置绘图区域
    
    def draw_plot(self):
        """ 对净值数据绘制净值走势曲线，对回撤数据绘制阴影图 """
        color_list_idx = 0 # 索引线条该取哪个颜色？
        for idx in range(self.n_cols - 1): # 第一列是日期，不能被包括进来
            series = self.chart.SeriesCollection(idx + 1)
            draw_type = AREA_PLOT if  "回撤" in series.Name else LINE_PLOT
            series.ChartType = draw_type # 绘图类型[只有回撤需要绘制面积图]
            if draw_type == LINE_PLOT:
                series.Format.Line.ForeColor.RGB = utils.RGB_tuple_to_float(utils.color_dict[self.line_color_list[color_list_idx]])
                series.Format.Line.Weight = 1.5 # Line 折线图粗细
                color_list_idx += 1
            if draw_type == AREA_PLOT: # 阴影图绘制颜色取默认的浅灰色
                series.AxisGroup = 2 # 设置次坐标轴
                series.Format.Fill.ForeColor.RGB = utils.RGB_tuple_to_float(utils.color_dict["shallow_grey"])
    
    def set_chart_style(self):
        """ 调整图表格式 """
        self.set_basic()
        self.set_legend()
        self.set_xaxis()
        self.set_yaxis("left")
        self.set_yaxis("right")
        self.set_grid()
    
    def set_basic(self):
        """ 基础设定：图表无填充，图表无边框 """
        self.chart.ChartArea.Format.Fill.Visible=False # 图片无填充
        self.chart.ChartArea.Format.Line.Visible=False #图片无边框
        self.chart.ChartArea.Format.Fill.Transparency = 1 
        self.chart.PlotArea.Interior.ColorIndex = False
    
    def set_legend(self):
        """ 设定图例选项 """
        # 标签置于顶部，其它数值参考 https://learn.microsoft.com/zh-cn/office/vba/api/excel.xllegendposition
        legend = self.chart.Legend
        legend.Position = -4160 
        legend.Font.Name = "楷体"
        legend.Font.Size = 11 # 图例字体大小
    
    def set_xaxis(self):
        """ 设置横坐标轴 """
        x_axis = self.chart.Axes(1) # 获取横轴对象
        x_axis.TickLabels.Font.Name = "Arial" # 设置坐标轴字体
        x_axis.TickLabels.Font.Size = 10 # 设置坐标轴字体大小
        x_axis.MajorUnit  = 3 # 设置坐标间隔
        x_axis.TickLabels.NumberFormat  = "yyyy-mm" # 设置数字格式
        x_axis.Border.Weight =  win32.constants.xlThin
        x_axis.Border.Color =  0x000000
        x_axis.MajorTickMark  =  win32.constants.xlTickMarkInside # 横轴刻度线向内
    
    def set_yaxis(self, _type: str = "left"):
        """
        设置左轴或者右轴

        Args:
            _type (str, optional): 有两个选项 [left, right]，分别表示设置左轴还是右轴.
        """
        if _type not in ["left", "right"]:
            raise ValueError(_type, "参数只能选择下列二者中的一个：", ["left", "right"])
        type_value = 1 if _type == "left" else 2 # 1代表左轴，2代表右轴
        y_axis = self.chart.Axes(2, type_value)
        y_axis.TickLabels.Font.Name = "Arial" # 设置坐标轴字体
        y_axis.TickLabels.Font.Size = 10 # 设置坐标轴大小
        y_axis.TickLabels.NumberFormat  = "0.00" if type_value == 1 else "0.0%" # 纵轴数值格式
        y_axis.Format.Line.Visible = False # 纵轴刻度线消失
    
    def set_grid(self):
        """ 设置网格线，默认和左轴刻度对齐 """
        y_axis = self.chart.Axes(2, 1)
        y_axis.HasMajorGridlines  = True # 给左轴设置灰色网格线
        y_axis.MajorGridlines.Border.Color = 0xD9D9D9 # 设置灰色网格线的颜色
        y_axis.MajorGridlines.Border.LineStyle = win32.constants.xlDash # 设置虚线风格

    def close_and_save(self):
        """ 关闭并保存得到的结果 """
        self.excel_file.Save()
        self.excel_file.Close()
        self.excel_app.Quit()


