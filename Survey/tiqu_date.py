import pandas
import re
import openpyxl

path="D:\\2022\\基坑监测\\华侨城6-1\\华侨城01数据库.xlsx"

DATA=pandas.read_excel(path,'日报')
book1=openpyxl.Workbook()
sheet1=book1.create_sheet('日报')
for i in range(DATA.shape[0]):
    sheet1.cell(i+1,1).value=DATA.iloc[i,1]
book1.save("D:\\2022\\基坑监测\\华侨城6-1\\tiqu.xlsx")