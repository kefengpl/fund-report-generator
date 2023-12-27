""" 此类用于处理 word 中的 Table 对象 """
import numpy as np
import win32com.client as win32

color_dict = {
    "deep_red" : 0xC00000,
    "orange" : 0xFFC000,
    "deep_grey" : 0x7F7F7F,
    "shallow_red" : 0xFFCCCC,
    "shallow_grey" : 0xBFBFBF
}
def convert_to_RGB(BGR_value: int):
    """ 通过位运算将 BGR 转为 RGB """
    blue = (BGR_value & 0xFF0000) >> 16
    green = (BGR_value & 0x00FF00) >> 8
    red = BGR_value & 0x0000FF
    return (red << 16) | (green << 8) | blue

class WordTableHandler:
    def __init__(self, table, title_mode: str, chinese_font: str = "楷体", 
                 english_font: str = "Times New Roman", font_size: float = 8):
        """
        此类用于处理 word table 对象

        Args:
            - table (_type_): 一个 word table 对象
            - title_mode (str): ["sep", "first", "first_row", "first_col"] 
                                "sep" 表示表头按行间隔，会填充第一列；从第一行开始间隔填充行;
                                "first_row" 表示表头只分布在第一行，所以只会对第一行进行填充;
                                "first_col" 表示表头只分布在第一列，所以只会对第一列进行填充;
                                "first" 表示表头分布在第一行以及第一列，所以会对第一行和第一列进行填充.
        """
        self.table = table
        self.table_fill_color = convert_to_RGB(color_dict["deep_red"]) # 填写16进制 RGB，表示表格需要填充的颜色
        self.row_height = 20 # 设置每行的行高
        self.chinese_font = chinese_font
        self.english_font = english_font
        self.font_size = font_size
        self.title_mode = title_mode
        if self.title_mode not in ["sep", "first", "first_row", "first_col"]:
            raise ValueError(r"title_mode 只能是(sep, first, first_row, first_col)之一")
    
    def get_rows(self) -> int:
        return len(self.table.Rows)

    def get_cols(self) -> int:
        return len(self.table.Columns)
    
    def set_cell_text_format(self, cell):
        """ 设置一个单元格的字体，对齐 """
        cell.Range.Font.Name = self.chinese_font
        cell.Range.Font.Name = self.english_font
        cell.Range.Font.Size = self.font_size
        cell.VerticalAlignment = win32.constants.wdCellAlignVerticalCenter # 设置垂直居中
        pf = cell.Range.ParagraphFormat # 获得单元格段落格式对象
        pf.Alignment = win32.constants.wdAlignParagraphCenter # 设置水平居中
        pf.LineSpacingRule = win32.constants.wdLineSpaceExactly # 设置固定行距
        pf.LineSpacing = 9 # 15 pond 1.25x 行距

    def fill_one_line(self, idx: int, fill_color: int, _type: str = "row"):
        """
        给某一行或者某一列上色

        Args:
            - table (_type_): word table 对象
            - idx (int): 给第几行上色？从0开始，0表示给第1行上色
            - fill_color (int): 填充颜色，应该是 RGB 值(16进制)，例如 0xFFEEFF
            - _type (str): ["row", "col"] 选择是给某一行上色，还是某一列上色，默认是行
        """
        if _type not in ["row", "col"]:
            raise ValueError("_type 的取值只能是 [row, col] 中的一种")
        line = self.table.Rows(idx + 1) if _type == "row" else self.table.Columns(idx + 1)
        line.Shading.BackgroundPatternColor = fill_color
    
    def is_table_header(self, row_idx: int, col_idx: int) -> bool:
        """
        查询某个单元格是否是表头单元格[即私募报告中被红色填充的位置]

        Args:
            - row_idx (int): 行索引，从0开始
            - col_idx (int): 列索引，从0开始
        """
        if self.title_mode == "first_col":
            return col_idx == 0
        if self.title_mode == "first_row":
            return row_idx == 0
        if self.title_mode == "first":
            return col_idx == 0 or row_idx == 0
        if self.title_mode == "sep":
            return col_idx == 0 or (row_idx % 2 == 0)

    def fill_table_color(self):
        """ 给表格的表头部分上色 """
        if self.title_mode == "first_col":
            self.fill_one_line(0, self.table_fill_color, "col")
        elif self.title_mode == "first_row":
            self.fill_one_line(0, self.table_fill_color, "row")
        elif self.title_mode == "first":
            self.fill_one_line(0, self.table_fill_color, "col")
            self.fill_one_line(0, self.table_fill_color, "row")
        else:
            self.fill_one_line(0, self.table_fill_color, "col")
            for idx in range(self.get_rows()):
                if idx % 2 == 0:
                    self.fill_one_line(idx, self.table_fill_color, "row")

    def set_rows_height(self):
        """ 为表格的每行设置统一的行高 """
        for row in self.table.Rows:
            row.Height = self.row_height
    
    def set_borders(self):
        """  为表格添加所有框线 """
        borders = [win32.constants.wdBorderTop, win32.constants.wdBorderLeft, 
                win32.constants.wdBorderBottom, win32.constants.wdBorderRight, 
                win32.constants.wdBorderHorizontal, win32.constants.wdBorderVertical]
        for border in borders:
            self.table.Borders(border).LineStyle = win32.constants.wdLineStyleSingle
            self.table.Borders(border).LineWidth = 6 

    def add_one_cell_text(self, row_idx: int, col_idx: int, text: str):
        """
        为一个单元格添加文本

        Args:
            - row_idx (int): 单元格在第几行，行索引从0开始
            - col_idx (int): 单元格在第几列，列索引从0开始
            - text (str): 该单元格中希望输入的文本
        """
        the_cell = self.table.Cell(row_idx + 1, col_idx + 1) # NOTE 注意 word VBA 中索引从1开始
        self.set_cell_text_format(the_cell) # 调整单元格文字格式
        the_cell.Range.Text = text # 输入单元格文本
        the_cell.Range.Font.Bold = True if self.is_table_header(row_idx, col_idx) else False  # 决定该单元格文字是否加粗

    def add_text(self, text_array: np.ndarray):
        """ 
        整体设置单元格内容

        Args:
            text_array (np.ndarray): 文本矩阵，矩阵中每个元素与表格位置一一对应，
                                     所以 text_array 和 word table 的行数和列数必须完全一致 

        Raises:
            ValueError: _description_
        """
        if (text_array.shape[0] != self.get_rows()) or (text_array.shape[1] != self.get_cols()):
            raise ValueError("输入是非法的，文本矩阵和word表格的形状必须完全一致")
        for row in range(self.get_rows()):
            for col in range(self.get_cols()):
                self.add_one_cell_text(row, col, str(text_array[row, col]))  