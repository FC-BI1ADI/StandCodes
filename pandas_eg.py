import pandas as pd

# NAN （数值数据类型的一类数），全称 Not a Number ，表示未定义或者不可表示的值。
# axis 坐标轴, ascending 升降序
# 创新一个dataframe
df = pd.DataFrame(np.arange(50).reshape(10, 5), columns=['a', 'b', 'c', 'd', 'e'])

# 打开EXCEL文件
xlsx_filename = r'EXCEL文件名.xlsx'
# sheet_name 表索引
# header 第N行为header, header行 pandas会将其作为列索引
# names 指定各列的列索引
# usecols 读入指定列
df = pd.read_excel(xlsx_filename, sheet_name=0, header=0, names=('A','B','C',..., 'X'), usecols=[0, 1, 2, ..., N])
# describe 统计下数据量，标准值，平均值，最大值等
df.describe()
# 调整各列的顺序或是在dataframe创建过程中用columns=[col_label]指定
order = [col_label1, col_label2, col_label3, ..., col_labelN]
df = df[order]
# 增加新列
df['col_label] = None
# 获取某列
col = df['col_label']
# 获取各列标签
col_label_list = df.columns.values
# 直接使用索引来选择dataframe数据，df还可以使用bool索引
df[s:e][col_label] # 左闭右开的规则
df[s:e][s:e] # 左闭右开的规则

# 以标签（行、列的名字）为索引选择数据—— x.loc[行标签,列标签]
# 以位置（第几行、第几列）为索引选择数据—— x.iloc[行位置,列位置]
# 同时根据标签和位置选择数据——x.ix[行,列]
# 需特别注意loc取得是标签，左闭右闭，而iloc取得是位置，左闭右开
x.loc[2:4] != x.iloc[2:4]
x[2:4] == x.iloc[2:4]

# pandas.DataFrame.values返回DataFrame的Numpy表示形式
df.values()
# 读取df的前几行,不指定参数默认为 5行
df.head(N)
# 读取df的后几行,不指定参数默认为 5行
df.tail(N)
# df数据的定位选择
df.loc[row_label]
df.iloc[row_index,col_index]


# 写入EXCEL文件
df.to_excel(xlsx_filename, sheet_name=0)

