import win32com.client

import openpyxl
import pandas as pd
path="D:\\Documents\\OneDrive\\工作\\2021\\基坑监测\\保利天珺A07地块\\日报\\保利天珺A07数据库.xlsx"
path1 = "D:\\Documents\\OneDrive\\工作\\2021\\基坑监测\\保利天珺A07地块\\质量评定.doc"
path2="D:\\Documents\\OneDrive\\工作\\2021\\基坑监测\\保利天珺A07地块\\质量评定\\"
import shutil
import os
import re
date=[]
data=pd.read_excel(path,'日报')
for i in range(data.shape[0]):
    date.append(data.iloc[i,1])
filenames=[]
date1=[]
for i in range(len(date)):
    x1 = str(date[i]).replace('-', '年', 1)
    x2 = x1.replace('-', '月', 1)
    x3 = x2.replace(' 00:00:00', '日', 1)
    shutil.copy(path1,path2+x3+'-'+'质量评定.doc')
    filenames.append(path2+x3+'-'+'质量评定.doc')
    date1.append(x3)
#传参，docx路径，需要替换的，替换的字
def replace_main(path,old_text,new_text):
    word = win32com.client.Dispatch("Word.Application") # 模拟打开 office
    doc = word.Documents.Open(path)
    word.Selection.Find.ClearFormatting()
    word.Selection.Find.Replacement.ClearFormatting()

    #1.True--区分大小写,2.True--完全匹配的单词，并非单词中的部分(全字匹配)3.True--使用通配符,
    # 4.True--同音,5.True--查找单词的各种形式,6.True--向文档尾部搜索,7.True--带格式的文本。
    # 2 - -替换个数(0表示不替换，1 表示只替换匹配到的第一个，2 表示全部替换)True--区分大小写,不可省略
    # word.Selection.Find.Execute(old_text, False, False, False, False, False, True, 1, False, new_text, 2)
    word.Selection.Find.Execute(old_text, False, False, False, False, False, True, 1, False, new_text, 2)
    doc.Close(SaveChanges=True)
    word.Quit()




if __name__ == '__main__':
    for i in range(len(date)):
        path1=filenames[i]
        year1 = re.findall(r'(.*)年', date1[i], flags=0)[0]
        month1 = str(int(re.findall(r'年(.*)月', date1[i], flags=0)[0]))
        day1 = str(int(re.findall(r'月(.*)日', date1[i], flags=0)[0]))
        old_text = "          年         月         日"
        # new_text = "      2022年    5月   12日"
        if(int(month1)<10 and int(day1)<10):
            new_text="      "+year1+'年'+'    '+month1+'月'+'    '+day1+'日'
        if(int(month1)<10 and (int(day1)>10 or int(day1)==10)):
            new_text="      "+year1+'年'+'    '+month1+'月'+'   '+day1+'日'
        if((int(month1)>10 or int(month1) == 10) and int(day1)<10):
            new_text="      "+year1+'年'+'   '+month1+'月'+'    '+day1+'日'
        if ((int(month1) > 10 or int(month1) == 10) and (int(day1)>10 or int(day1)==10)):
            new_text="      "+year1+'年'+'   '+month1+'月'+'   '+day1+'日'
        replace_main(path1,old_text,new_text)