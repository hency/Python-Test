from matplotlib.patches import Ellipse, Circle
import matplotlib.pyplot as plt
import random

# 产生 1 到 10 的一个整数型随机数
# for i in range(10):
#     print(random.randitn(1, 10))
# print(random.randint(1, 10))
# # 产生 0 到 1 之间的随机浮点数
# print(random.random())
# # 产生  1.1 到 5.4 之间的随机浮点数，区间可以不是整数
# print(random.uniform(1.1, 5.4))
# # 从序列中随机选取一个元素
# print(random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9, 0]))
# # 生成从1到100的间隔为2的随机整数
# print(random.randrange(1, 100, 2))
# # 将序列a中的元素顺序打乱
# a = [1, 3, 5, 6, 7]
# import numpy as np
#
# # 产生n维的均匀分布的随机数
# print(np.random.rand(5, 5, 5))
#
# # 产生n维的正态分布的随机数
# print(np.random.randn(5, 5, 5))
#
# # 产生n--m之间的k个整数
# print(np.random.randint(1, 50, 5))
#
# # 产生n个0--1之间的随机数
# print(np.random.random(10))
#
# # 从序列中选择数据
# print(np.random.choice([2, 5, 7, 8, 9, 11, 3]))
# numx2=[]
# if(len(numx1)>1):
#     print("*******************%d****************************************"%len(numx1))
# else:
#     numx3=[]
#     flag=numx1[0]
#     index1=0
#     for i in range(1,len(numx1)):
#         if(numx1[i]-flag==i-index1):
#             print("*********************************************连续**************************")
#             numx3.append(numx1[i])
#         else:
#             print('*********************************************断开*************************')
#             flag=numx1[i]
#             index1=i
numx1=[1,2,5,7,8,11,22,23,24]
flag=0
numx2=[]
for i in range(1,len(numx1)):
    if(numx1[i]-numx1[flag]==i-flag):
        if(i==len(numx1)-1):
            numx2.append(numx1[flag:i+1])
        pass
    else:
        if(i-flag==1):
            numx2.append(numx1[flag])
            if(i==len(numx1)-1):
                numx2.append(numx1[i])
        else:
            numx2.append(numx1[flag:flag+i-flag])
        flag=i
