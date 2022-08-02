import pandas
import openpyxl
import re
import os
path_dataset="D:\\2022\\基坑监测\\二附院数据库\\二附院数据库20220502.xlsx"
data=pandas.read_excel(path_dataset,'日报')
date=[]
date_num=[]
for i in range(data.shape[0]):
    x1 = str(data.iloc[i,1]).replace('-', '年', 1)
    x2 = x1.replace('-', '月', 1)
    x3 = x2.replace(' 00:00:00', '日', 1)
    date.append(x3)
    date_num.append(data.iloc[i,0])
path="D:\\2022\\基坑监测\\二附院数据库\\质量评定\\"
filenames=os.listdir(path)
for i in range(len(filenames)):
    date1=re.findall(r'(.*)-质量评定',filenames[i],flags=0)[0]
    for z in range(len(date)):
        if(date1==date[z]):
            # os.rename(path + filenames[i],
            #           path + re.findall(r'(.*)第',filenames[i],flags=0)[0]+'.doc')
            os.rename(path + filenames[i],
                      path + re.findall(r'(.*)[.]',filenames[i],flags=0)[0]+'第'+str(date_num[z])+'期'+'.doc')




