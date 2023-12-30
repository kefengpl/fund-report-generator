""" 此类用于处理 Word """
import os
import win32com.client as win32
import numpy as np
import utils

import word_table_handler as wth
import excel_chart_handler as ech

class WordHandler:
    def __init__(self, path: str = None, visible: bool = True, 
                 chinese_font: str = "楷体", english_font: str = "Times New Roman"):
        """
        此类之作用在于操作 WORD 生成文字、段落、表格。

        Args:
            - path (str): 你希望写入word文档的绝对路径，该 word 文档需要是空文档。如果该参数是 None，将自动创建新文档
            - visible (bool, optional): 是否希望在写入 word 的过程中打开word并跟踪这一过程? 默认值 True.
            - chinese_font (str, optional): 中文字体设定. Defaults to "楷体".
            - english_font (str, optional): 英文字体设定. Defaults to "Times New Roman".
        """
        self.word_app = win32.gencache.EnsureDispatch('Word.Application') # 创建 WORD APP 应用程序
        self.word_app.Visible = visible
        self.is_new_doc: bool = True if path is None else False # 此变量用于最后保存时进行判断
        # 如果给了空文档路径，就打开这个空文档，否则就创建新文档
        self.this_doc = self.word_app.Documents.Open(path) if path is not None else self.word_app.Documents.Add()   
        self.cursor = self.word_app.Selection # 获取光标(也可以代表选择区域)对象
        self.chinese_font = chinese_font
        self.english_font = english_font
        self.font_size = { # 字体在这里设定即可
            "title" : 12, # 一级标题：12号字[小四]
            "paragraph" :  12, # 二级标题：10.5号字[五号]
            "footnote" : 10.5 # 脚注字体大小[9号字小五]
        }
        self.line_spacing = { # 设置行间距，起始有换算公式，例如：1.25倍行距 = 字体大小 * 1.25
            "title" : self.font_size["title"] * 1.5, # 18 pond
            "paragraph" : self.font_size["paragraph"] * 5 / 3,  # 20 pond
            "footnote" : self.font_size["footnote"] * 1.25, # 1.25 倍行距
        }

    def set_page_layout(self):
        """ 设置页面边距，注意：72磅是1英寸是2.54cm """
        # 设置页面边距，单位为磅（1 英寸 = 72 磅）
        top_margin = 72    # 1 inch
        bottom_margin = 72 # 1 inch
        left_margin = 72   # 1 inch
        right_margin = 72  # 1 inch

        # 设置页面尺寸，单位为磅（1 英寸 = 72 磅，1 厘米 ≈ 28.35 磅）
        #page_height =  (27.94 / 2.54) * 72  # 假设页面高度为27.94cm
        #page_width = (21.59 / 2.54) * 72  # 假设页面宽度为21.59cm

        # 获取新文档的页面设置对象，并调整页面边距
        page_setup = self.this_doc.PageSetup
        page_setup.TopMargin = top_margin
        page_setup.BottomMargin = bottom_margin
        page_setup.LeftMargin = left_margin
        page_setup.RightMargin = right_margin
        page_setup.Orientation = win32.constants.wdOrientLandscape 

        # 获取新文档的页面设置对象，并调整页面尺寸
        #page_setup.PaperSize = win32.constants.wdPaperCustom
        #page_setup.PageHeight = page_height
        #page_setup.PageWidth = page_width
    
    def set_font_format(self, chinese_font: str, english_font: str, font_size: int, bold: bool):
        """
        设置字体格式

        Args:
            - chinese_font (str): 中文字体
            - english_font (str): 英文与数字字体
            - font_size (int): 字号
            - bold (bool): 是否加粗
        """
        self.cursor.Font.Name = chinese_font
        self.cursor.Font.Name = english_font
        self.cursor.Font.Size = font_size
        self.cursor.Font.Bold = bold
    
    def set_paragraph_format(self, line_spacing: float, alignment: int = 0):
        """
        设置行距和对齐

        Args:
            - line_spacing (float): 行距是多少磅？换算公式，例如：1.25倍行距 = (字体大小 * 1.25) 磅
            - alignment (int): 对齐模式；0是左对齐，4是 真 · 两端对齐
        """
        # 获取选中光标所在段落，调整该段落格式
        paragraph_format = self.cursor.ParagraphFormat
        # paragraph_format.LineSpacingRule = win32.constants.wdLineSpaceExactly # 设置固定行距
        paragraph_format.LineSpacingRule = win32.constants.wdLineSpaceAtLeast # 设置最小行距
        paragraph_format.LineSpacing = line_spacing # 15 pond 1.25x 行距
        paragraph_format.Alignment = alignment # 4 是两端对齐
        return None
    
    def cursor_move_down(self):
        """ 在添加完成一段("一段"可以是一个标题或者一个段落)后，取消选中该段落，然后使得光标下移 """
        # 取消选中，然后下移
        self.cursor.Collapse(Direction=win32.constants.wdCollapseEnd)
        self.cursor.Text = "\r" # 新建立一行
        self.cursor.EndKey(Unit=win32.constants.wdStory) # 到达文章尾部
    
    def add_text_content(self, text: str, _type: str = "paragraph"):
        """
        添加标题 或者段落，一级标题最大，二级标题较小。当然，目前来看，不需要用到 level 这些参数。\n
        流程很简单：设置字体 --> 添加文本 --> 调整段落 --> 光标下移

        Args:
            - text (str): 添加的文字内容，例如 "1. 中金量化-银河海山 1 号【CTA 日内短线高频策略】"
            - type (str): 添加一段还是一个标题？ paragraph 表示添加一个段落， title 表示添加标题，footnote 表示添加脚注[就是图表左下方的文字]
        """
        if _type not in ["title", "paragraph", "footnote"]:
            raise ValueError(_type, "参数必须是下列值之一：", ["title", "paragraph", "footnote"])
        text = text.strip()
        text = "    " + text if _type == "paragraph" else text
        self.set_font_format(self.chinese_font, self.english_font, 
                             self.font_size[_type], True if _type == "title" else False)
        self.cursor.Text = text
        self.set_paragraph_format(self.line_spacing[_type])
        self.cursor_move_down()    

    def add_picture(self, picture_path: str):
        """ 添加图片，暂时只能居左 """
        this_paragraph = self.cursor.Paragraphs.Add()
        # NOTE 插入图片的时候必须设置段落是最小值多少磅，否则图像插入将异常
        this_paragraph.Format.LineSpacingRule = win32.constants.wdLineSpaceAtLeast
        this_paragraph_range = this_paragraph.Range
        #指定文件的完整路径
        picture_path = picture_path 
        #在当前的段落中插入图片
        this_paragraph_range.InlineShapes.AddPicture(picture_path)
        # 取消选中并且下移
        self.cursor.MoveDown(win32.constants.wdParagraph, 1)
    
    def add_excel_chart(self, file_path: str):
        """ 绘制EXCEL图表到WORD里，需要指定EXCEL数据文件路径 """
        # NOTE 插入图片的时候必须设置段落是最小值多少磅，否则图像插入将异常
        self.cursor.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceAtLeast
        chart_handler = ech.ExcelChartHandler(file_path, False)
        chart_handler.draw_plot()
        chart_handler.set_chart_style()
        chart_handler.chart.ChartArea.Copy() # 复制图片
        self.cursor.Paste() # 粘贴图片
        self.cursor_move_down()
        chart_handler.close_and_save()
    
    def add_table(self, n_rows: int, n_cols: int, title_mode: str, text_array: np.ndarray):
        this_table = self.cursor.Tables.Add(self.cursor.Range, NumRows = n_rows, NumColumns = n_cols)
        # 获取页面信息
        page_setup = self.this_doc.PageSetup
        page_width = page_setup.PageWidth
        left_margin = page_setup.LeftMargin
        right_margin = page_setup.RightMargin
        table_width = page_width - (left_margin + right_margin)
        # 处理表格格式，并添加文本
        table_handler = wth.WordTableHandler(this_table, table_width, title_mode)
        table_handler.fill_table_color()
        table_handler.set_borders()
        table_handler.set_rows_height()
        table_handler.add_text(text_array)
        self.cursor.EndKey(Unit=win32.constants.wdStory)
    
    def close_and_save(self, fund_name: str):
        """ 关闭并保存得到的结果 """
        if self.is_new_doc:
            self.this_doc.SaveAs(os.path.abspath("output/" + utils.generate_filename(fund_name, ".docx")))
        else:
            self.this_doc.Save()
        self.this_doc.Close()
        self.word_app.Quit()