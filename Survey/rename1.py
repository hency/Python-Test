import os
path="D:\\Desktop\\赣电中心1\\"
path1="D:\Desktop\赣电中心\第2期2020年10月23日\质量评定.doc"
name1=os.listdir(path)
import os
def SearchAbsPath(dirname):
    dirname = os.path.abspath(dirname)
    filenames = list()
    for root,dirs,files in os.walk(dirname, topdown=False): #扫描一层目录
        for name in files:
            filenames.append(root+os.path.sep+name) #每一个文件的绝对路径放入列表
            if(name[0:4]=='周边道路'):
                p1=root+os.path.sep+name
                try:
                    times=int(name[-6:-4])
                    p2=root+os.path.sep+'周边道路、周边管线沉降'+str(times)+'.dat'
                    os.rename(p1,p2)
                    print(p2)
                except Exception as f:
                    times=int(name[-5])
                    p2=root+os.path.sep+'周边道路、周边管线沉降'+str(times)+'.dat'
                    os.rename(p1,p2)
                    print(p2)
                finally:
                    print('')
    return filenames
def SearchAbsPath1(dirname1):
    dirname = os.path.abspath(dirname1)
    filenames = list()
    for root,dirs,files in os.walk(dirname, topdown=False): #扫描一层目录 topdown指定优先于Top目录或者优先锁定子目录
        for name in files:
            filenames.append(root+os.path.sep+name) #每一个文件的绝对路径放入列表
            if(name[0:4]=='周边建筑'):
                p1=root+os.path.sep+name
                try:
                    times=int(name[-6:-4])
                    p2=root+os.path.sep+'周边建筑沉降'+str(times)+'.dat'
                    os.rename(p1,p2)
                    print(p2)
                except Exception as f:
                    times=int(name[-5])
                    p2=root+os.path.sep+'周边建筑沉降'+str(times)+'.dat'
                    os.rename(p1,p2)
                    print(p2)
                finally:
                    print('')
    return filenames
SearchAbsPath(path)
SearchAbsPath1(path)

