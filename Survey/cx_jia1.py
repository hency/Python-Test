import pandas
import openpyxl
import re
import os
path_dataset="D:\\2022\\基坑监测\\大剧院\\数据库\\保利大剧院数据库1 -还原原始数据用的.xlsx"
data=pandas.read_excel(path_dataset,'日报')
date=[]
date_num=[]
for i in range(data.shape[0]):
    x1 = str(data.iloc[i,1]).replace('-', '年', 1)
    x2 = x1.replace('-', '月', 1)
    x3 = x2.replace(' 00:00:00', '日', 1)
    date.append(x3)
    date_num.append(data.iloc[i,0])
path="D:\\2022\\基坑监测\\大剧院\\数据库\\原始数据存放位置\\测斜\\"
filenames=os.listdir(path)
for i in range(len(filenames)):
    filenames1=os.listdir(path+filenames[i])
    for j in range(len(filenames1)):
        # os.rename(path+filenames[i]+'\\'+filenames1[j],path+filenames[i]+'\\'+filenames[i]+'-'+filenames1[j])
            date1='20'+re.findall(r'-(.*)00时',filenames1[j],flags=0)[0]
            for z in range(len(date)):
                if(date1==date[z]):
                    os.rename(path + filenames[i] + '\\' + filenames1[j],
                              path + filenames[i] + '\\' + re.findall(r'(.*).xlsx',filenames1[j],flags=0)[0]+'第'+str(date_num[z])+'期'+'.xlsx')




