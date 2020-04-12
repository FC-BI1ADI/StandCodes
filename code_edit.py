# 本文档主要定义Python代码编辑中的通用规则

# 变量定义
# 循环变量的定义：常用单字母表示，如 i,j,k
for i in range(10):
    print(i)
# --------------------
# 计数器
count = 0
count += 1
# --------------------
# 最大值 max, 最小值 min
max = 10
min = 0
# --------------------
# python中的字符串一般情况均只使用单引号'string',除特殊情况下可使用双引号"string"
str = 'string'
# --------------------
# python的常用数据结构
# 列表 list， 元组 tuple, 集合 set, 字典 dictionary dict
# 以上数据结构一般使用下划线+结构的方式表达
data_list = []
data_tuple = ()
data_set = {}
data_dict = {key1:value1, key2:value2, ...}
# --------------------
# 与日期时间相关变量定义
# 日期 date： year 年， month 月， day 日
# date, year, mon, day
# 时间 time: hour 小时， minute 分钟， second 秒
# time, hour, min, sec
# 定时器 timer
# --------------------
# 与文件相关变量定义
# 路径 path
# 文件 file, filename
# 目录 directory， dir
path = 'd:\\temp\test.xlsx'
dir = 'd:\\temp'
file = 'test.xlsx'
filename = 'test.xlsx'
# --------------------
# 与表格相关变量定义
# row 表示行， rows 表示多行
# col 表示列， cols 表示多列, column
# index 表示索引
# per_ 表示每XX
for per_row in table.rows:
    for per_col in table.cols:
        pass
for row_index in range(0, tables_rows + 1):
    for col_index in range(0, tables_cols + 1):
        print(table[row_index, col_index])
# --------------------
