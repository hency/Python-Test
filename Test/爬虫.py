import requests #网络请求
import re #正则模块
import time

for i in range(1,167):
    a_url = "http://www.dytt8.net/html/gndy/dyzz/list_23_"+str(i)+".html"
    html = requests.get(a_url)
    html.encoding = "gb2312" #指定编码
    detil_list = re.findall('<a href="(.*?)" class="ulink">',html.text)
    for j in detil_list:
      #  time.sleep(2)
        b_url = "http://www.dytt8.net"+j #电影完整网址
        html_2 = requests.get(b_url)
        html_2.encoding = "gb2312"
        ftp = re.findall('<a href="(.*?)">.*?</a></td>',html_2.text)
        try:
            with open(r"D:\DYTT\LJ.txt","a",encoding = "utf-8") as file: 
                file.write(ftp[0]+"\n")
        except:
            print(b_url+"这一页没有")