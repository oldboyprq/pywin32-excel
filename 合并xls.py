from win32com.client.gencache import EnsureDispatch
import os

excel = EnsureDispatch('Excel.Application')
dir_path = os.path.join(os.getcwd(), 'all.xlsx')
book = excel.Workbooks.Open(dir_path)
print(book.Name)
for root, dirs, files in os.walk(os.getcwd()):
    for file in files:  # 现有文件名
        if '.xls' == file[-4:]:
            book1 = excel.Workbooks.Open(os.path.join(os.getcwd(), file))
            print(book1.Name)
            # 一个excel内有多张表此处可以加一个循环
            book1.Sheets(1).Select()
            book1.Sheets(1).Copy(After=book.Sheets(1))
            #
            book1.Close()
book.Save()
book.Close()
excel.Quit()
