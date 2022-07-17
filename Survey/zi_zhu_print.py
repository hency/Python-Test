import win32print
import tempfile
import win32api
import os
path="D:\\Desktop\\汇总"
file_names=os.listdir(path)
file_names1=file_names[10:]
def print_file(filename):
    open(filename,"r")
    win32api.ShellExecute(
        0,
        "print",
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )
##进行排序file_names 泡沫排序法
for i in range(len(file_names)):
    if(isinstance(file_names[i][1],int)):
        pass
    else:
        pass

# for i in range(0,5):
#     print_file(path+'\\'+file_names1[i])