import os
import re
path="D:\\2022\\吕建勋生成原始数据\\坡顶转桩顶\\"
filenames=os.listdir(path)
for i in range(len(filenames)):
    os.rename(path+filenames[i],path+'桩顶'+re.findall('坡顶(.*)',filenames[i],flags=0)[0])
