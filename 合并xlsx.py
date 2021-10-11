from win32com.client.gencache import EnsureDispatch
import os

excel = EnsureDispatch('Excel.Application')
dir_path = os.getcwd()
book = excel.Workbooks.Open(os.path.join(dir_path, 'all.xlsx'))
print(book.Name)
for root, dirs, files in os.walk(dir_path):
    for file in files:  # 现有文件名
        if ('all' not in file) and ('~$' not in file) and ('.xlsx' == file[-5:]):
            book1 = excel.Workbooks.Open(os.path.join(dir_path, file))
            print(book1.Name)
            # 一个excel内有多张表此处可以加一个循环
            book1.Sheets(1).Select()
            book1.Sheets(1).Copy(After=book.Sheets(1))
            #
            book1.Close()
book.Save()
book.Close()
excel.Quit()
