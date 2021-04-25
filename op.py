# 对excel进行操作

from win32com.client import Dispatch  # pip install pywin32
import os

path = os.path.join(os.getcwd(), 'test.xlsx')  # excel文件路径
excel = Dispatch('Excel.Application')
book = excel.Workbooks.Open(path)
sheet = book.Worksheets('Sheet1')

# 对score小于60的name进行标红
# 第一种方法(遍历所有的行）
# rows = sheet.UsedRange.Rows.Count   # sheet已用的行
# for i in range(2, rows + 1):
#     if sheet.Cells(i, 4).Value < 60:
#         sheet.Cells(i, 2).Font.Color = -16776961      # 红色，数值见录制宏后的vba代码

sheet.AutoFilterMode = False  # 非筛选模式
# 第二种方法（excel自带的筛选，然后选中筛选结果，改变选中部分字体颜色，或者对筛选的结果进行遍历，逐行改变）
sheet.Range("$A:$E").AutoFilter(Field := 4, Criteria1 := "<60")
# 对筛选出的结果进行遍历，逐行改变
# for i in sheet.AutoFilter.Range.SpecialCells(12):
#     if i.Row > 1:                   # 第一行为表头
#         sheet.Cells(i.Row, 2).Font.Color = -16776961
# sheet.AutoFilterMode = False
# 这里不必关注筛选出了哪些行，直接选中数据的第一行至最后一行即可,其实选中的就是筛选出的部分。
sheet.Range("$B2:$B{}".format(sheet.UsedRange.Rows.Count)).Select()
excel.Selection.Font.Color = -16776961

# 筛选出各课程得分最高的人，name列字体颜色改为蓝色，加粗
# 读取C列和D列的值得到各最大的值，然后对B列进行操作（逐行遍历、筛选）
sheet.AutoFilterMode = False
row = sheet.UsedRange.Rows.Count
name_score_max = dict()
for i in range(2, row + 1):
    x = sheet.Cells(i, 3).Value
    name_score_max[x] = max(sheet.Cells(i, 4).Value, name_score_max.get(x, 0))
# 逐行遍历
# for i in range(2, row + 1):
#     x = sheet.Cells(i, 3).Value
#     if name_score_max[x] == sheet.Cells(i, 4).Value:
#         sheet.Cells(i, 2).Font.Color = -1003520
#         sheet.Cells(i, 2).Font.Bold = True
# 筛选操作
for i in name_score_max.keys():
    sheet.AutoFilterMode = False
    sheet.Range("$A:$E").AutoFilter(Field := 3, Criteria1 := "={}".format(int(i)))
    sheet.Range("$A:$E").AutoFilter(Field := 4, Criteria1 := "={}".format(int(name_score_max[i])))
    sheet.Range("$B2:$B{}".format(row)).Select()
    excel.Selection.Font.Color = -1003520
    excel.Selection.Font.Bold = True
sheet.AutoFilterMode = False
