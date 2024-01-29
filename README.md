## 自动化项目工程文件

如何使用该文件？

**准备数据(这里是导出单个基金产品的报告的数据示例，本文件下方给出了批量导出时的数据格式)** 

两个 xlsx 文件：1.某只基金的净值数据 2. 该基金希望对标的指数数据

- 合理的基金净值数据文件示例如下。说明：①基金净值列的列名需要和基金名称相同，净值开始值可以不是1.00；②最左侧必须是日期列，第二列必须是净值列。③虽然给出了第三列，但是第三列是无效的(在调用导出单个基金报告接口的情况下，如果是批量导出，则第三列不会忽略)，在程序执行过程中会被直接忽略。最好只给出前两列数据即可。

| |	九章幻方 |	000905.SH | 
| ---- | ---- | ---- | 
| 2017-01-20 |	1.00 |	6,121.9983| 
|2017-01-26 |	1.00 	|6,223.7061 |
|2017-02-03	|1.00| 	6,207.0921|
|2017-02-10|	1.00 |	6,337.1081|
|2017-02-17|	1.00| 	6,307.1627|
|2017-02-24	|1.00 	|6,476.1655|
|2017-03-03|	1.00| 	6,452.8390|
|2017-03-10|	1.01 |	6,447.9165|
|2017-03-17	|1.02| 	6,483.2463|

- 合理的指数数据文件示例如下。说明：①指数价格走势列名必须是指数名称。②最左侧必须是日期列。③对于非指增基金，后面几列可以包含若干指数数据，这些指数数据都会被绘制在净值走势图中；对于指增基金，只能包含一列指数数据，即该指增基金对标的指数数据。

|	       | 中证1000 | 中证500 |
| ---- | ---- | ---- |
|2022-09-02 | 6776.8268   |  12345 |
|2022-09-09	| 6913.579    |  67891 |
|2022-09-16	| 6481.2442   |  51515 |
|2022-09-23	| 6363.6835   |  17194 |
|2022-09-30	| 6124.8639   |  12548 |
|2022-10-14	| 6406.4573   |  12846 |
|2022-10-21	| 6432.3255   |  18460 |
|2022-10-28	| 6239.9316   |  89131 |
|2022-11-04	| 6709.6446   |  19841 |

**设置参数** 在main.py文件分别设置两个xlsx文件数据的路径，指定是否是指增基金，指定私募管理人名称，下面的代码是一个示例。

```
def single_fund_report_interface():
    """ 获得基金分析报告 """
    # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略(如果是批量导出函数，则不会忽略后面的数据)。
    netval_path = "data/裕锦中证1000指数增强-净值数据.xlsx"
    # 指数数据表路径：① 非指增：允许添加多列指数数据；② 指增：只允许添加一列指数数据 
    index_path = "data/裕锦中证1000指数增强-指数数据.xlsx" 
    # 需要在参数中手动指明是否是指增类基金，False表示不是指增基金，True表示指增基金
    enhanced_fund = True
    # 私募基金管理人名称，不想写可以删了
    corp_name = "裕锦量化" 
     # 指定开始计算的日期，当然，该参数可以删除，这行也可以删除或者赋值为 None，默认从净值数据的起始日期开始计算。
    start_date = date(2022, 8, 19) 
    single_fund_report(netval_path, index_path, enhanced_fund, corp_name, start_date = start_date)

def single_fund_indicator_tables_interface():
    """ 生成单个基金产品各类指标(不包括月度/年度指标，因为它们已经汇总在 fund_report中了)汇总表，滚动收益统计表，放在word中 """
    netval_path = "data/图灵进取中证1000指数增强-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    corp_name = "图灵量化" # 私募基金管理人名称，不想写可以删了
    single_fund_indicator_tables(netval_path, corp_name)  

def multi_fund_report_interface():
    """ 在基金对的标指数数据一致的情况下，允许同时导出多个基金的基金报告。"""
    netval_path = "data/中证1000基金-净值数据.xlsx" # 净值数据表路径，可以包含多个基金，每个基金的数据是一列，列名是基金名称。不要出现任何冗余列。
    index_path = "data/裕锦中证1000指数增强-指数数据.xlsx"  # 所有基金共用该指数数据表。指数数据表路径：① 非指增：允许添加多列指数数据；② 指增：只允许添加一列指数数据
    enhanced_fund = True # 指定是否为指增基金：注意：指增基金和非指增基金不能同时批量计算
    start_dates = [] # 从净值数据从前往后按顺序指定起始计算日期。如果有三个基金，仅指定前两个日期，可以写为 [date(2022, 8, 19), date(2022, 9, 19)]
                     # 如果有三个基金，需要指定第一个基金和第三个基金的开始计算日期，可以写为 [date(2022, 8, 19), None, date(2022, 9, 19)]
    corp_names = ["裕锦私募", "图灵量化"] # 从前往后按顺序匹配公司名称，使用方法与 start_dates 的使用完全一致
    add_indicators_tables: bool = True # 添加 “关键指标汇总”, “滚动收益率分布”, “收益概率统计” 这三张表
    multi_fund_report(netval_path, index_path, enhanced_fund, corp_names = corp_names, 
                      start_dates = start_dates, add_indicators_tables = add_indicators_tables)

def multi_fund_indicator_tables_interface():
    """ 生成多个基金产品各类指标(不包括月度/年度指标，因为它们已经汇总在 fund_report中了)汇总表，滚动收益统计表，放在word中 """
    netval_path = "data/中证1000基金-净值数据.xlsx" # 净值数据表路径，只会处理前两列数据，后面列的数据会被直接忽略。
    corp_names: list = ["裕锦私募", "图灵量化"] # 私募基金管理人名称列表，不想写可以删了
    start_dates: list = [] # 可以指定每只基金的起始计算日期，必须是date格式，例如 date(2022, 8, 19)
    multi_fund_indicator_tables(netval_path, corp_names = corp_names, start_dates = start_dates)  

def main():
    single_fund_report_interface()
```

除了上述接口函数之外，还实现了多个接口，用以实现不同功能，这些接口如下，详情和使用方法需要参照[main.py](main.py)。

如果使用某个接口，需要在接口函数中填好参数，然后在 main() 中调用该接口函数即可。

```
single_fund_report_interface() # 导出单个基金的产品报告
single_fund_indicator_tables_interface() # 导出单个基金的关键指标及滚动收益统计
multi_fund_report_interface() # 导出多个基金的产品报告
multi_fund_indicator_tables_interface() # 导出多个基金的关键指标及滚动收益统计
```

日期设置。Python 中有多种日期格式，为了保证统一处理，在代码中手动输入的日期格式必须是 datetime.date 格式，下面是一个示例。

```
from datetime import date
start_date: date = date(2023, 8, 19) # 注意：日期需要这样按照 年-月-日 来创建
```

**其它功能**

程序是十分灵活的，在编写时预设了很多可对外调用的函数和方法。如果不希望导出WORD，只希望根据净值数据计算一些指标，比如夏普、年化等，可以参考 [interactive_code.ipynb](interactive_code.ipynb)。

**测试程序时的运行环境**
```
- python 3.8.8
- numpy 1.24.4
- pandas 2.0.3
- pywin32 306
```

**更新情况**

**2024-01-26** 添加新功能，除了导出基金报告的WORD外，还可以导出关键指标和滚动收益分析的WORD。

**2024-01-26** 支持有限度的基金报告批量导出。“有限度”指的是

- 基金需要对齐时间序列，比如以表格最左侧的时间为基准，放入同一个excel文件里面，且每列列名是基金名，不允许有冗余的列。
下面是一个数据表的示例，允许每列开头的净值数据是空值。

| |	九章幻方 |	乐子壬固收1号 | 
| ---- | ---- | ---- | 
| 2017-01-20 |	1.00 |	| 
|2017-01-26 |	1.00 	|  |
|2017-02-03	|1.00| 	1.00|
|2017-02-10|	1.00 |	1.07|
|2017-02-17|	1.00| 	1.08|
|2017-02-24	|1.00 	|1.09|
|2017-03-03|	1.00| 	1.10|
|2017-03-10|	1.01 |	1.11|
|2017-03-17	|1.02| 	1.02|

- 对标的指数必须一致，因为只允许输入一个指数数据文件，这些基金共享这个指数数据文件的所有指数数据。
例如：①同为1000指增的10个基金可以将净值数据放入1张 excel ②同时希望对标两个指数 [中证转债，中证债券指数] 的8个非指增基金的净值数据可以放入 1 张 excel。

- 由上一条引申出：批量生成报告必须使得所有基金要么都是指增基金，要么都是非指增基金。
  