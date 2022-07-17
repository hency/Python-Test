import os
import shutil
import re
#获取指定文件中文件名
def get_filename(filetype):
    name =[]
    final_name_list = []
    source_dir=os.getcwd()#读取当前路径
    for root,dirs,files in os.walk(source_dir):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    final_name_list = [item +filetype for item in name]
    return final_name_list #返回由文件名组成的列表
#筛选文件，利用正则表达式
def select_file(str_cond,file_name_list):
    select_name_list =[]
    part1 = re.compile(str_cond)#正则表达式筛选条件
    for file_name in file_name_list:
        if len(part1.findall(file_name)):#判断其中一个文件名是否满足正则表达式的筛选条件
            select_name_list.append(file_name)#满足，则加入列表
    return select_name_list#返回由满足条件的文件名组成的列表
#复制指定文件到另一个文件夹里，并删除原文件夹中的文件
def cope_file(select_file_name_list,old_path,new_path):
    for file_name in select_file_name_list:
        shutil.copyfile(os.path.join(old_path,file_name),os.path.join(new_path,file_name))#路径拼接要用os.path.join，复制指定文件到另一个文件夹里
        os.remove(os.path.join(old_path,file_name))#删除原文件夹中的指定文件文件
    return select_file_name_list
#主函数
def main_function(filetype,str_cond,old_path,new_path):
    final_name_list = get_filename(filetype)
    select_file_name_list = select_file(str_cond,final_name_list)
    cope_file(select_file_name_list,old_path,new_path)
    return select_file_name_list

file_type = '.csv'#指定文件类型
str_cond = '-Dfn_info-'#正则条件
old_path = 'F:\\data\\text_1'#原文件夹路径
new_path = 'F:\\data\\dfn_info'#新文件夹路径
main_function(file_type,str_cond,old_path,new_path)#主函数