"""
Convert Excel files to PDF
"""
import pandas
import glob
import openpyxl


# 将多个文件路径读入list:filepaths
filepaths = glob.glob("excels/*.xlsx")
# print(filepaths)

for filepath in filepaths:
    content = pandas.read_excel(filepath, sheet_name="Sheet 1")
    print(content)






