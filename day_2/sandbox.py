import os
import openpyxl
from openpyxl.workbook.workbook import Workbook

wb = openpyxl.load_workbook('day_2/lecture.xlsx')
# print(wb.workbook)

# wb = Workbook()
# print(str(wb))

for i in range(3, 10):
  wb.create_sheet()

for sheet in wb:
  y = 5
  sheet.title = 'MyNewSheet' + str(y)
  y += 1

print(str(wb.sheetnames))