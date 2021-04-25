# 生成随机数据
import random
from win32com.client import Dispatch
import string
import os

path = os.path.join(os.getcwd(), 'test.xlsx')
excel = Dispatch('Excel.Application')
book = excel.Workbooks.Open(path)
sheet = book.Worksheets('Sheet1')
letter = string.ascii_lowercase


def name():
    names = ""
    for _ in range(random.randint(3, 5)):
        names = names + random.choice(letter)
    return names


# excel 表头 id	name	class	score	gender
for i in range(1, 21):
    sheet.Cells(i + 1, 1).Value = i
    sheet.Cells(i + 1, 2).Value = name()
    sheet.Cells(i + 1, 3).FormulaR1C1 = "=RANDBETWEEN(1,5)"    # 使用excel中的公式Randbetween()
    sheet.Cells(i + 1, 4).FormulaR1C1 = "=RANDBETWEEN(40,100)"
    sheet.Cells(i + 1, 5).FormulaR1C1 = "=RANDBETWEEN(0,1)"


book.Save()
book.Close()
excel.Quit()
