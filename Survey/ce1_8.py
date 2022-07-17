# import numpy
# import pandas as pd
# import numpy as np
# import random
# from matplotlib import pyplot
# import openpyxl
# from openpyxl import load_workbook
# from openpyxl.styles import Font
# from 沉降观测1 import read_measure_line_from_dataset
# # path1='C:\\Users\\Kong\\Desktop\\项目\\测试期次1\\output1.xlsx'
# path1="D:\\Desktop\\测试期次1\\output1.xlsx"
# # path2='C:\\Users\\Kong\\Desktop\\项目\\测试期次1\\output2.xlsx'
# path2='D:\\Desktop\\测试期次1\\output2.xlsx'
# # dateset_path='C:\\Users\\Kong\\Desktop\\项目\\测试期次1\\测试数据库1.xlsx'
# dateset_path='D:\\Desktop\\测试期次1\\测试数据库1.xlsx'
# # measure_line_path="C:\\Users\\Kong\\Desktop\\项目\\测试期次1\\测试线路文件1.xlsx"
# measure_line_path='D:\\Desktop\\测试期次1\\测试线路文件1.xlsx'
# RZ_names,RZ_values=read_measure_line_from_dataset(dateset_path,measure_line_path)
# def mid(x,y):
#     a=x[0]+(y[0]-x[0])/2
#     b=x[1]+(y[1]-x[1])/2
#     return [a,b]
# # path="C:\\Users\\Kong\\Desktop\\项目\\测试期次1\\测试线路文件1.xlsx"
# path="D:\\Desktop\\测试期次1\\测试线路文件1.xlsx"
# df=pd.read_excel(path,'Sheet1')#对应的导出为pd.to_excel(***.xlsx)
# df.columns=range(0,3)
# name=RZ_names
# value=RZ_values[0]
# dict = {name[i]:value[i] for i in range(len(name))}
# for i in range(0,df.shape[0]):
#     if(df.loc[i,0] in name):
#         df.loc[i,1]=dict[df.loc[i,0]]
# df1=df[:1]
# df2=df[1:len(df)]
# df21=df2.dropna(how='any')
# df3=pd.concat([df1,df21],axis=0)
# df3.index=range(df3.shape[0])
# a=df3.iloc[:,[1]]
# b=range(df3.shape[0])
# num=[]
# numx=[]
# for i in range(df3.shape[0]):
#     if(df3.loc[i,0] in name):
#         numx.append(i)
#         pass
#     else:
#         num.append(i)
# df4=df3.loc[num,[0,1,2]]
# position=position_index(df3)
# # BC1=0.010
# numx=[2,3]
# num1=[0,1,4,5]
# dh1=df3.loc[position[0][1],1]-df3.loc[position[0][0],1]
# # Hz=df3.loc[1,1]-1*BC1-dh1
# # SHZ=SHB+dh1
# SJ=df3.loc[position[0][0],2]
# Z_name=df3.loc[position[0][0],0]

# sheet1.cell(7 + 5 * i, 9).value = height1
def position_index1(df3):
    num = []
    position1 = position_index(df3)
    flag=0
    k1=0
    k3=0
    while(flag==0):
        for q1 in range(k1,len(position1)):
            k=0
            k2=position1[q1][1]
            for q2 in range(len(position1)):
                if(position1[q1][1]==position1[q2][1]):
                    k=k+1
                    k1=k1+1
                else:
                    pass
            k3=k3+k
            num.append([k2,k,k3])
            break
        if(k1==len(position1)):
            flag=1
    return num
# num_1=position_index1(df3)
# for zz1 in range(len(position)):
#     HD_random3 = round(random.uniform(-0.5, 0.5), 3)  ##单位m
#     Z_name.append(df3.loc[position[zz1][0], 0])
#     Z_HD.append(df3.loc[position[zz1][0], 2] + HD_random3)  ###
#     Z_DH1 = df3.loc[dict1[position[zz1][1]], 1] - df3.loc[position[zz1][0], 1]
#     Z_H.append(df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - Z_DH1)
#     Z_HB = RH1_random[position[zz1][1]]
#     Z_HF.append(Z_HB + Z_DH1)
#     if (Z_HB + df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - 1.75 < Z_H[zz1] and Z_H[zz1] < Z_HB +
#             df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - 0.6):
#         print("支点的高程设置合理！！！！！！！！！！！")
#     else:
#         print("支点的高程设置不合理****************")
def Z_H_function(df3,df4,num,BC1):
    dict1 = {i: num[i] for i in range(df4.shape[0])}
    position=position_index(df3)
    Z_name = []
    Z_H = []
    Z_HD = []
    Z_HF = []
    for zz1 in range(len(position)):
        pass
        Z_DH1 = df3.loc[dict1[position[zz1][1]], 1] - df3.loc[position[zz1][0], 1]
        Z_H.append(df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - Z_DH1)
    return Z_H
def position_index(df3):
    position = []
    num_1=0
    for i in range(df3.shape[0]):
        if 'Y' in df3.iloc[i,0]:
            position.append([i,i-num_1-1])
            num_1 = num_1 + 1
    return position
# Z_H=Z_H_function(df3,df4,num,BC1)
    # if (Z_HB + df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - 1.75 < Z_H[zz1] and Z_H[zz1] < Z_HB +
    #         df3.loc[dict1[position[zz1][1]], 1] - position[zz1][1] * BC1 - 0.6):
    #     print("支点的高程设置合理！！！！！！！！！！！")
    # else:
    #     print("支点的高程设置不合理****************")
