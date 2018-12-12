# df2xlsx

An easy way to export DataFrame objects as tables and charts to excel file.

方便、简洁地将Pandas的DataFrame地输出为Excel的表格及图表,

## Installation 安装

```
pip install https://github.com/UncleJiong/pdexcel.git
```

## Examples

简单的例子:

```python
   import numpy as np
   import pandas as pd
   from pd2xlsx import *
   
   dfa = pd.DataFrame(np.random.rand(14, 2), columns=list('AB'))
   dfb = pd.DataFrame(np.random.rand(10, 4), columns=list('ABCD'))

   # 创建Excel文件
   writer = ExcelWriter('demo2.xlsx')

   # 创建工作表Sheet1
   sheet1 = writer.add_sheet(sheetname='Sheet1')
   # 添加表格1
   table11 = sheet1.add_table(dfa, table_name='Table_1')
   # 添加表格2, 数值格式百分比两位小数，列宽设置为8.
   table12 = sheet1.add_table(dfb, dicformat={'num_format':'0.00%'}, width=6)
   # 插入图表, 默认为折线图
   chart11 = table11.add_chart()
   # 插入直方图, 标题'Column Chart'
   chart12 = table12.add_chart(chart_name='Column Chart', chart_type='column')
   # 插入 A、C两列数据的折线图, 高度为默认的2倍
   chart13 = table12.add_chart(chart_col=['A','C'], y_scale=2)
   
   # 创建工作表Sheet2
   sheet2 = writer.add_sheet()
   # 插入指定风格的表格
   table2 = sheet2.add_table(dfa, tbl_style='Table Style Light 11')
   # 插入指定风格的图表
   chart2 = table2.add_chart(chart_style=37)
   
   # 退出并保存文件
   writer.close()
```

<div align="center">
  <img src="https://raw.github.com/UncleJiong/pdexcel/master/example/demo1a.png"><br><br>
</div>

<div align="center">
  <img src="https://raw.github.com/UncleJiong/pdexcel/master/example/demo1b.png"><br><br>
</div>


也可以通过`to_excel`函数更简洁地生成Excel文件:

```python
   import numpy as np
   import pandas as pd
   from pd2xlsx import *
   
   dfa = pd.DataFrame(np.random.rand(14, 2), columns=list('AB'))
   dfb = pd.DataFrame(np.random.rand(10, 4), columns=list('ABCD'))

   with ExcelWriter('demo.xlsx') as writer2:
       # 表格插入Sheet1, 对A、C字段数据绘折线图
       sheet1 = to_excel(writer2, dfa, kwargs_chart=dict(chart_col=['A', 'C']))
       # 表格插入Sheet1, 不绘图, 数值以百分比格式保存
       sheet1 = to_excel(writer2, dfb, sheet1, chart=False,
                         kwargs_cell={'num_format':'0.00%'})
       # 表格插入Sheet2, 绘直方图
       sheet2 = to_excel(writer2, dfa, chart_type='column')
```
	   

<div align="center">
  <img src="https://raw.github.com/UncleJiong/pdexcel/master/example/demo2a.png"><br><br>
</div>

<div align="center">
  <img src="https://raw.github.com/UncleJiong/pdexcel/master/example/demo2b.png"><br><br>
</div>