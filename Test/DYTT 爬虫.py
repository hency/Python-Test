import requests,re,time

for i in range(1,167):
    url_a = 'http://www.dytt8.net/html/gndy/dyzz/list_23_'+str(i)+'.html'
    html_a = requests.get(url_a)
    html_a.encoding = 'gb2312'
    detil_list = re.findall('<a class="ulink" href="(.*?)">',html_a.text)
    for j in detil_list:
        url_b = 'http://www.dytt8.net'+detil_list[j]
        html_b = requests.get(url_b)
        html_b.encoding = 'gb2312'
        ftp = re.findall('<a href="(.*?)">.*?</a></td>',html_b.text)
        try:
            with open(r'D:\DYTT.txt','a',encoding='utf-8') as file:
                file.write(ftp[0]+'\n')
        except:
            print(url_b + "这一页没有匹配到")