path="D:\\2022\\吴昌程生成原始数据\\1\\最终报告\\道路"

import os
import re
names=os.listdir(path)
for name in names:
    os.rename(path+'\\'+name,path+'\\'+'周边道路'+re.findall(r'周边地表(.*)',name,flags=0)[0])