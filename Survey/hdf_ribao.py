import re
import os
import numpy as np
import shutil
path="D:\\Desktop\\黄地发还原原始数据\\金地原始数据\\"
path1="D:\\Desktop\\黄地发还原原始数据\\日报\\"
name1=os.listdir(path)
for i in range(len(name1)):
    name2=os.listdir(path+name1[i]+'\\')
    flag=0
    for j in range(len(name2)):
        if('日报' in name2[j]):
            flag=1
            break
    if(flag==1):
        print('不用新建文件夹：日报')
        if(len(os.listdir(path+name1[i]+'\\'+'日报'+'\\'))==1):
            pass
        else:
            print('拷贝对应的日报过来')
            name3=os.listdir(path1)
            for z in range(len(name3)):
                if(name3[z][-4:]=='xlsx'):
                    if(int(re.findall('(.*)监测日报',name3[z],flags=0)[0]) == int(re.findall('第(.*)期',name1[i],flags=0)[0])):
                        shutil.copy(path1+name3[z],path+name1[i]+'\\'+'日报')
    else:
        os.makedirs(path+name1[i]+'\\'+'日报')
        name3 = os.listdir(path1)
        for z in range(len(name3)):
            if (name3[z][-4:] == 'xlsx'):
                if (int(re.findall('(.*)监测日报', name3[z], flags=0)[0]) == int(
                        re.findall('第(.*)期', name1[i], flags=0)[0])):
                    shutil.copy(path1 + name3[z], path + name1[i] + '\\' + '日报')
