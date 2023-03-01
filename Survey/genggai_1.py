import re
import os
path="D:\\2022\\基坑监测\\华侨城4-2\\更改1"
filename=os.listdir(path)
for i in range(len(filename)):
    os.rename(path+'\\'+filename[i],path+'\\'+'坡顶沉降'+re.findall('件(.*)',filename[i],flags=0)[0])