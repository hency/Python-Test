import os
import re
path="D:\\2022\\基坑监测\\二附院数据库\\测斜"
filenames=os.listdir(path)
for i in range(len(filenames)):
    filenames1=os.listdir(path+'\\'+filenames[i])
    for j in range(len(filenames1)):
        os.rename(path+'\\'+filenames[i]+'\\'+filenames1[j],path+'\\'+filenames[i]+'\\'+filenames[i]+'-'+filenames1[j])
