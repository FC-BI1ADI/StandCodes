import openpyxl_eg
from openpyxl_eg import Workbook
from openpyxl_eg import load_workbook

# openpyxl中索引编号均是从1开始，注意与pandas区别

# 读入EXCEL文件
excel_filename = 'd:\\data\\test.xlsx'
wb = openpyxl_eg.load_workbook(excel_filename)
ws = wb.active
# 获取EXCEL文件中所有表
sheet_list = openpyxl_eg.workbook.Workbook.sheetnames()
# 设置表的标题和按标题获取表
ws.title = 'new title'
new_ws = wb['new title']
# 设置表的标题颜色
ws.sheet_properties.tabColor = "RRGGBB"
# openpyxl允许直接遍历表 You can loop through worksheets
for sheet in wb:
    print(sheet.title)
# 访问ws中的单元格,允许使用标号或切片
cell_value = ws['A4']
ws['A4'] = 4
ws.cell(row=4, column=2, value=10)
ws.cell(row_index,col_index)
cell_range = ws['A1':'C2']
colC = ws['C']
col_range = ws['C:D']
row10 = ws[10]
row_range = ws[5:10]
# 遍历ws中的行和列，基于1的索引表达
# 参数：min_row=None, max_row=None, min_col=None, max_col=None
for per_row in ws.iter_rows(min_row=1, max_col=3, max_row=2):
    pass
for col in ws.iter_cols(min_row=1, max_col=3, max_row=2):
    pass

# 写入EXCEL文件
wb = Workbook()
ws = wb.active
dest_filename = 'empty_book.xlsx'
# 获取活动表单或创建表单
ws1 = wb.active
ws2 = wb.create_sheet(title='sheet2')
ws3 = wb.create_sheet(title='sheet3')
# 写入EXCEL文件
wb.save(filename = dest_filename)

