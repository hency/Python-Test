import pandas
import numpy
import openpyxl
from openpyxl import load_workbook
import os
import re
import shutil
import xlsxwriter
from openpyxl.drawing.image import Image
from win32com.client import constants
#跳过打开excel表格的步骤，sht代表要处理的表


import win32com.client



picture_path1="D:\\2022\\基坑监测\\电子签名png格式\\曾荡电子签.png"
picture_path2="D:\\2022\\基坑监测\\电子签名png格式\\吴昌程电子签.png"
picture_path3="D:\\2022\\基坑监测\\电子签名png格式\\黄地发电子签.png"

path="D:\\2022\\基坑监测\\科创大街\\科创大街数据库-报表\\17监测日报2022-4-18.xlsx"

excel = win32com.client.DispatchEx('Excel.Application')  # 这个是必备的，
#使用win32建新excel也需要他
excel.Visible = True  # 是否可视化
wb = excel.Workbooks.Open(path, UpdateLinks=False, ReadOnly=False)
ws=wb.Worksheets['封面']
for shp in ws.Shapes:
    if(shp.Type == constants.msoPicture):
        shp.Delete() #z只删除图片的shape
#
####worksheet.insert_image(row=1, col=5, filename='E:\\临时图片\\' + v, options={'x_scale': 0.1, 'y_scale': 0.1})

image=Image(picture_path1)
image1=Image(picture_path2)
image2=Image(picture_path3)
data=pandas.read_excel(path,'封面')
book1=load_workbook(path)
sheet1=book1.get_sheet_by_name('封面')
image.width=140
image.height=120
image1.width=140
image.height=120
image2.width=100
image2.height=70

sheet1.add_image(image,'B8')
sheet1.add_image(image1,'B9')
sheet1.add_image(image2,'B10')
book1.save(path)
