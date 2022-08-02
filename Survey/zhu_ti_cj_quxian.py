# -*- coding: UTF-8 -*-
import matplotlib
from matplotlib import pyplot
# 1.首先导包
from matplotlib import pyplot
# 2. 准备数据：坐标轴的刻度数据以及构成图的数据
# 2.1 首先准备坐标轴的刻度数据
from matplotlib.font_manager import FontProperties  # 导入FontProperties
font = FontProperties(fname="SimHei.ttf", size=14)  # 设置字体
labelX=[1,2,3,4,5]
labelY=[-5,-4,-3,-2,-1,0,5,10,30]
# 2.2 再准备构成图的数据
# 第一个折线图的数据
x1=['2020/6/6'	,'2020/6/7'	,'2020/7/29'	,'2020/9/6',	'2020/10/9'	,'2020/11/5'	,'2020/11/24'	,'2020/12/17',	'2021/1/14',	'2021/3/9'	,'2021/3/31',	'2021/5/18',	'2021/8/18',	'2021/11/18',	'2022/2/18'	,'2022/6/10']
y1=[0,0	,-1.1	,-2.1	,-3	,-3.7	,-4.1,	-4.4,	-5	,-5.8,	-6.5,	-7.1,	-7.5,	-7.7,	-7.81	,-7.81]
# 第二个折线图的数据
x2=['2020/6/6'	,'2020/6/7'	,'2020/7/29'	,'2020/9/6',	'2020/10/9'	,'2020/11/5'	,'2020/11/24'	,'2020/12/17',	'2021/1/14',	'2021/3/9'	,'2021/3/31',	'2021/5/18',	'2021/8/18',	'2021/11/18',	'2022/2/18'	,'2022/6/10']
y2=[1,3,5,7,9,11,13,15,17,19,21,23,25,27,27,27]
# 3.然后准备画布，决定图的宽、高、清晰度(20是宽，8是高，dpi是清晰度)
pyplot.figure(figsize=(20,8),dpi=80)
# 4.将构成图的数据绑定到图上，先是横坐标，然后是纵坐标，label是标记，标价显示还需要legend()
# 画第一个折线图
pyplot.plot(x1, y1,label="第一个折线图")
# 画第二个折线图,自动改变颜色！当然也可以指定两个折线图分别为什么颜色
pyplot.plot(x2, y2,label="第二个折线图")
# 显示线的标记
pyplot.yticks(labelY, size = 12)
# Python matplotlib画图时图例说明(legend)放到图像外侧详解：https://www.jb51.net/article/186659.htm
pyplot.legend(loc=7,prop=font)
# 5.将坐标轴刻度绑定上去，然后再标记x、y分别代表了什么;刻度和标记都为Times New Roman，且字体大小为16
# 5.1 绑定刻度，刻度数据可以通过一一对应显示字符串
# # 5.2 标记x、y代表什么
pyplot.xlabel("代表x轴",fontdict={'weight' : 1000,'size' : 16},fontproperties=font,rotation=45)
pyplot.ylabel("代表y轴",fontdict={'weight' : 'normal','size' : 16},fontproperties=font)
ax=pyplot.gca()
ax.xaxis.set_ticks_position('top')    # 将x轴刻度设在下面的坐标轴上
ax.yaxis.set_ticks_position('left')         # 将y轴刻度设在左边的坐标轴上
ax.spines['top'].set_position(('data', 0))   # 将两个坐标轴的位置设在数据点原点
ax.spines['left'].set_position(('data', 0))
# 6.整个图的标题
pyplot.title("图的标题",fontproperties=font)
# 7.背景换成网格，以及添加水印
# 7.1 网格：ls=":"-->网格样式（虚线）,color="gray"-->网格颜色,alpha=0.5-->网格透明度
pyplot.grid(ls=":",color="gray",alpha=0.5)
# 7.2 添加水印
pyplot.text(x=1,               # 水印开头左下角对应的X点
 		 y=2,               # 水印开头左下角对应的Y点
         s="Matplotlib",    # 水印文本
         fontsize=50,       # 水印大小
         color="gray",      # 水印颜色
         alpha=0.5)         # 水印是通过透明度控制的
pyplot.xticks(rotation='vertical')
# 8.保存图
pyplot.savefig("./save.png")
# 9.显示图
pyplot.show()


