from win32com.client.gencache import EnsureDispatch
import os

# 各个sheet需汇总合并的字段
input_col = list(input("输入想要的字段，以英文空格分隔\n").split(' '))
excel = EnsureDispatch('Excel.Application')
cur_path = os.path.join(os.getcwd(), 'all.xlsx')
book = excel.Workbooks.Open(cur_path)
sheet_num = book.Worksheets.Count
sheet_all = excel.Worksheets('all')
# 填充表头
for i in range(len(input_col)):
    sheet_all.Cells(1, i + 1).Value = input_col[i]

for i in range(sheet_num):
    # 得到汇总使用行
    cur_row = sheet_all.UsedRange.Rows.Count + 1
    sheet = book.Worksheets(i + 1)
    if sheet.Name == 'all':
        continue
    print(sheet.Name)
    row = sheet.UsedRange.Rows.Count
    col = sheet.UsedRange.Columns.Count
    # print(row,col)
    # 在每一个sheet表头中查找字段对应列
    index_list = [0] * len(input_col)
    for k in range(col):
        if sheet.Cells(1, k + 1).Value in input_col:
            # 匹配到的是表头中的第几个，对应下标记录
            input_col_index = input_col.index(sheet.Cells(1, k + 1).Value)
            index_list[input_col_index] = k + 1
        if 0 not in index_list:
            break

    for j in index_list:  # 各sheet对应all表头下标
        try:
            sheet.Activate()
            sheet.Range(sheet.Cells(2, j), sheet.Cells(row, j)).Select()
            excel.Selection.Copy()
            sheet_all.Activate()
            sheet_all.Cells(cur_row, index_list.index(j) + 1).Select()
            sheet_all.Paste()
        except Exception as e:
            print(e)
book.Save()
book.Close()
excel.Quit()

# 故障大类，完成处理部门，故障号，工单流水号，请人工核查
