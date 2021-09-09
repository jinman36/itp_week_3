import os
import openpyxl

workbook = openpyxl.load_workbook('day_2/lecture.xlsx')
print(str(workbook))