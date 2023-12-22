# In[0]
import xlwings as xw
from openpyxl.utils import get_column_letter as Int2Col
from openpyxl.utils import column_index_from_string as Col2Int

#转换excel中的RGB值为xlwings中对应的BGR值
def RGB_BGR(color):
    RGB=color.split(',')
    BGR=RGB[2]+RGB[1]+RGB[0]
    return BGR

app=xw.App(visible=True,add_book=False)
#wb=app.books.open(r"D:\中信建投\中行策略报告\测试.xlsx")
wb=xw.Book(r"测试.xlsx")
ws=wb.sheets[0] # 取第0张表
nrow=ws.used_range.rows.count #获取数据区域行数
ncon=ws.used_range.columns.count #获取数据区域列数
firstrow="A1:"+Int2Col(ncon)+"1" #设置第一行区域：Int2Col [2 --> A 6 --> F]
firstcon="A1:A"+str(nrow) #设置第一列区域
maxrange=Int2Col(ncon)+":"+Int2Col(ncon) #设置最后一列区域

ws.api.UsedRange.Font.Name="Arial" #设置数据区域字体
ws.api.UsedRange.Borders.LineStyle=1 #设置数据区域边框
ws.api.UsedRange.Borders.Weight=2 #设置数据区域边框
ws.range("A:A").Font.Name = "Arial"
ws.api.UsedRange.HorizontalAlignment=-4108 #设置数据区域水平居中
ws.api.UsedRange.VerticalAlignment=-4108 #设置数据区域垂直居中
ws.api.UsedRange.ColumnWidth=11 #设置数据区域列宽
ws.range('A:A').number_format="yyyy-m-d" #设置第一列日期格式
ws.range(maxrange).number_format="0.00%" #设置最后一列回撤数据格式
ws.range(firstcon).color=(192,0,0) #设置第一列单元格颜色
ws.range(firstcon).api.Font.Color=0xFFFFFF #设置第一列字体颜色
ws.range(firstrow).color=(192,0,0) #设置第一行单元格颜色
ws.range(firstrow).api.Font.Color=0xFFFFFF #设置第一行字体颜色
wb.save()

chart=ws.charts.add()
chart.set_source_data(ws.range('A1').expand()) #设置图片数据区域
chart.left=400 #图片左边距
chart.top=10 #图片上边距
chart.width=400 #图片宽度
chart.height=200 #图片高度
chart.api[1].SetElement(104) #设置图例位于下方
chart.api[1].ChartArea.Format.Fill.Visible=False #图片无填充
chart.api[1].ChartArea.Format.Line.Visible=False #图片无边框
chart.api[1].ChartArea.Format.TextFrame2.TextRange.Font.NameComplexScript="楷体" #设置图片字体
chart.api[1].ChartArea.Format.TextFrame2.TextRange.Font.NameFarEast="楷体" #设置图片字体
chart.api[1].ChartArea.Format.TextFrame2.TextRange.Font.Name="楷体" #设置图片字体
chart.api[1].PlotArea.Format.Fill.Visible=False #绘图区无填充
chart.api[1].PlotArea.Format.Line.Visible=False #绘图区无边框
chart.api[1].Axes(1).TickLabelPosition=-4134 #横轴标签置底
chart.api[1].Axes(1).MajorUnit=6 #横轴标签间隔
chart.api[1].Axes(1).MajorUnitScale=1 #横轴标签间隔类型
chart.api[1].Axes(1).TickLabels.NumberFormat="yyyy/m" #横轴数据格式
chart.api[1].Axes(2,1).Format.Line.Visible=False #主纵轴无实线
chart.api[1].Axes(2,1).MajorUnit=0.05 #主纵轴标签间隔
chart.api[1].Axes(2,1).MaximumScale=1.15 #主纵轴标签最大值
chart.api[1].Axes(2,1).MinimumScale=0.95 #主纵轴标签最小值
chart.api[1].Axes(2,1).TickLabels.NumberFormat="0.00" #主纵轴数据格式
chart.api[1].Axes(2).MajorGridlines.Format.Line.DashStyle=4 #主网格线线形

#循环画图
for i in range(ncon-1):
    chart.api[1].SeriesCollection(i+1).ChartType=4

# chart.api[1].SeriesCollection(1).ChartType=4 #设置系列1图片类型
# color1="192,0,0"
# chart.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB=RGB_BGR(color1) #设置系列1颜色
# chart.api[1].SeriesCollection(2).ChartType=4
# chart.api[1].SeriesCollection(2).Format.Line.DashStyle=4 #设置系列2线形
# chart.api[1].SeriesCollection(3).ChartType=4
# chart.api[1].SeriesCollection(3).Format.Line.Weight=1.5 #设置系列3粗细
# chart.api[1].SeriesCollection(4).ChartType=4
# chart.api[1].SeriesCollection(5).ChartType=1
# chart.api[1].SeriesCollection(5).AxisGroup=2 #设置系列5为副纵轴
# chart.api[1].Axes(2,2).Format.Line.Visible=False #副纵轴无实线

wb.save()
wb.close()
app.quit()

"""
# In[1]
import win32com.client as win32
from PIL import ImageGrab

def transparence2white(img):
    sp=img.size
    width=sp[0]
    height=sp[1]
    print(sp)
    for yh in range(height):
        for xw in range(width):
            dot=(xw,yh)
            color_d=img.getpixel(dot)  # 与cv2不同的是，这里需要用getpixel方法来获取维度数据
            if(color_d[3]==0):
                color_d=(255,255,255,255)
                img.putpixel(dot,color_d)  # 赋值的方法是通过putpixel
    return img

excel=win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible=True #显示excel程序
wb=excel.Workbooks.Open(r"D:\中信建投\中行策略报告\测试\测试.xlsx")
ws=wb.Worksheets('Sheet1')
word=win32.gencache.EnsureDispatch('Word.Application')
word.Visible=True #显示word程序
doc=word.Documents.Open(r"D:\中信建投\中行策略报告\测试\测试.docx")
ws.UsedRange.Copy() #复制Excel数据区域
table_range=doc.Range()
#table_range.InsertAfter('\r\n') #粘贴表格前，先在word现有内容最后换行
table_range.Collapse(0) #设置从word现有内容最后开始粘贴
table_range.Paste()
doc.Tables(1).AutoFitBehavior(1) #表格根据内容自动调整

ws.UsedRange.CopyPicture() #复制数据区域为图片
ws.Paste() #粘贴图片
ws.Shapes(ws.Shapes.Count).Copy() #复制图片
img=ImageGrab.grabclipboard().convert('RGBA') #读取剪切板图片并转格式
img=transparence2white(img) #调整图片背景
img.convert('RGB').save(r"D:\中信建投\中行策略报告\测试\表格.jpg") #保存图片

i=len(ws.Shapes)
while i>=1:
    shape=ws.Shapes.Item(i)
    shape.Copy()
    shape_range=doc.Range()
    shape_range.InsertAfter('\r\n') #粘贴图片前，先在word现有内容最后换行
    shape_range.Collapse(0) #设置从word现有内容最后开始粘贴
    shape_range.Paste()
    i=i-1

doc.Save()
#doc.Close()
#word.Quit() #关闭退出word
excel.Quit() #关闭退出excel
"""
