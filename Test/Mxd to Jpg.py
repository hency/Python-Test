#-*- coding: UTF-8 -*-

##########################################################################################
#Author：mtl
#用途：将一个文件夹中的所有mxd文件(包括所有子目录中的mxd文件)导出成jpg图片，并放到各自的目录下。
#用法：用记事本打开，将参数path更改成指定的文件路径，注意将反斜杠“\”改成正斜杠“/”；res是设定的dpi值。
#运行方法：打开Python2.6中的IDLE，File - >Open -> run -> run Module。
##########################################################################################
import arcpy, os, time

#存放mxd文件的目录，也可以是存放mxd文件的上一级目录。
path = 'C:/Users/zhency/Desktop/000000000000000'
#导出jpg文件的分辨率。
res = 300
#mode1可选值为0或1，0表示导出path这个目录及其所有层次子目录中的mxd，1表示只导出path这个目录的mxd文件。
mode1 = 1
#mode2可选值为0或1，0表示mxd导出的图片放到与mxd相同的文件夹下，1表示放到path下面。
mode2 = 1
#mode3可选值为0或1，0表示导图结束后不关机，1表示结束后关机。
mode3 = 0

def main():
    for root, dirs, files in os.walk(path):
      if mode2 == 0:
        temp_path = root
      else:
        temp_path = path
      for afile in files:
        if afile[-4:].lower() == '.mxd':
          mxd = arcpy.mapping.MapDocument(os.path.join(root,afile))
          arcpy.mapping.ExportToJPEG(mxd, os.path.join(temp_path,afile[:-3] + 'jpg'), resolution = res)
          del mxd
          print ur'succeed in exporting file ' + afile[:-3] + ur'jpg'
        if mode1 != 0:
            break
    if mode3 != 0:
        os.system('shutdown -s -t 120')

if __name__ == "__main__":
    main()
