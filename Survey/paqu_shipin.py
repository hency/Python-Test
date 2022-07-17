import requests        #导入requests库
from bs4 import BeautifulSoup  #引入BeautifulSoup库
import time
from tqdm import  tqdm

# 查看章节列表信息
#
# 引入BeautifulSoup对网页内容进行解析
#
# 获取网页电子书文本信息



def get_content(target):

    req = requests.get(url=target,verify=False)  # 发起请求，获取html信息

    req.encoding = 'utf-8'  # 设置编码

    html = req.text  # 将网页的html信息保存在html变量中

    bf = BeautifulSoup(html, 'lxml')  # 使用lxml对网页信息进行解析

    texts = bf.find('div', id='content')  # 获取所有<div id = "content">的内容

    content = texts.text.strip().split('xa0' * 4)

    return content


if __name__ == '__main__':          #主函数入口

    server = 'https://www.xsbiquge.com'     #电子书网站地址

    book_name = '《元尊》.txt'

    # target = 'https://www.xsbiquge.com/78_78513/'#要爬取的目标地址,《元尊》的章节目录网址
    target='https://www.xsbiquge.com'

    req = requests.get(url=target)      #发起请求，获取html信息

    req.encoding='utf-8'                #设置编码

    html = req.text                     #将网页的html信息保存在html变量中

    chapter_bs = BeautifulSoup(html,'lxml')     #使用lxml对网页信息进行解析

    chapters = chapter_bs.find('div',id='list') #获取所有<div id = "list">的内容

    chapters = chapters.find_all('a')         #找到list中的a标签中的内容

    for chapter in tqdm(chapters):

        chapter_name = chapter.string           #章节名字

        url = server + chapter.get('href')       #获取章节链接中的href

        content = get_content(url)

    with open(book_name,'a',encoding='utf-8') as f:

        f.write("《"+chapter_name+"》")

        f.write('n')

        f.write('n'.join(content))

        f.write('n')