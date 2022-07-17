import re
import os
###需要保证原始数据文件夹下至少有一个数据文件夹才能保证len(os.listdir())>0
pathx1="D:\\Desktop\\黄地发还原原始数据\\金地\\"
pathx2=os.listdir(pathx1)
for j in range(len(os.listdir(pathx1))):
    pathx3=pathx2[j]
    path2=pathx1+pathx3+'\\原始数据\\'
    name1=pathx3
    a=re.findall(r'第(.*)期',name1,flags=0)[0] ####第128期2021.10.16格式为这个格式
    # pathx2='D:\\Desktop\\黄地发还原原始数据\\第128期2021.10.16\\原始数据\\坡顶沉降\\'+'水准观测文件'+a+'.dat'
    # path2='D:\\Desktop\\黄地发还原原始数据\\'+name1+'\\原始数据\\'
    b=re.findall('期(.*)[.]',name1,flags=0)[0][0:4]
    c=re.findall(r'[.](.*)[.]',name1,flags=0)[0]
    d=re.findall(r'[.](.*)',re.findall(r'[.](.*)',name1,flags=0)[0],flags=0)[0]
    if(len(c)<2):
        c='0'+c
    if (len(d) < 2):
        d = '0' + d

    str1='期次127 日期：2021/10/15'
    str1='期次'+a+' 日期：'+b+'/'+c+'/'+d
    str2=b+c+d+'.dat'
    str3='127观测记录2021-10-15坐标'
    str3=a+'观测记录'+b+'-'+c+'-'+d+'坐标'+'.dat'
    strc='坡顶沉降' ##文价夹为这个
    strw='坡顶位移' ##文价夹为这个
    strd='周边道路' ##文价夹为这个
    xmname1=os.listdir(path2)
    for i in range(len(xmname1)):
        if(xmname1[i]==strc):
            path3=path2+strc
            if(len(os.listdir(path3+'\\'))==1):
                path4=path3+'\\'+os.listdir(path3+'\\')[0]
                n = 0
                with open(path4) as f1:
                    message = ''
                    for line in f1:
                        if (n == 0):
                            line = str1 + '\n'
                        if (n == 1):
                            line = re.findall('(.*)T', line, flags=0)[
                                       0] + 'TO  ' + str2 + '               |                      |                      |                      |' + '\n'
                        message += line
                        n = n + 1
                with open(path3+'\\'+'水准观测文件'+a+'.dat', 'w') as f2:
                    f2.write(message)
                os.remove(path4)
        if(xmname1[i]==strw):
            path3=path2+strw
            if(len(os.listdir(path3+'\\'))==1):
                path4=path3+'\\'+os.listdir(path3+'\\')[0]
                with open(path4) as f1:
                    message = ''
                    for line in f1:
                        message += line
                with open(path3+'\\'+str3, 'w') as f2:
                    f2.write(message)
                os.remove(path4)
        if(xmname1[i]==strd):
            path3=path2+strd
            if(len(os.listdir(path3+'\\'))==1):
                path4=path3+'\\'+os.listdir(path3+'\\')[0]
                n = 0
                with open(path4) as f1:
                    message = ''
                    for line in f1:
                        if (n == 0):
                            line = str1 + '\n'
                        if (n == 1):
                            line = re.findall('(.*)T', line, flags=0)[
                                       0] + 'TO  ' + str2 + '               |                      |                      |                      |' + '\n'
                        message += line
                        n = n + 1
                with open(path3+'\\'+'水准观测文件'+a+'.dat', 'w') as f2:
                    f2.write(message)
                os.remove(path4)
    # path1="D:\Desktop\黄地发还原原始数据\第128期2021.10.16\原始数据\坡顶沉降\水准观测文件127.dat"
