# -*- coding:utf-8 -*-
import xlwt
import re
import requests
from bs4 import BeautifulSoup
import time

book_urls = []
booknum = 0
params = [{'start': 0}, {'start': 25}, {'start': 50}, {'start': 75}, {'start': 100},
          {'start': 125}, {'start': 150}, {'start': 175}, {'start': 200}, {'start': 225}, ]
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.6,en;q=0.4',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.93 Safari/537.36'
}

# get url of every book
rerule = re.compile('href="(.+?/\d+?/)"')
for start in params:
    res = requests.get('https://book.douban.com/tag/散文', params=start, headers=headers)
    book_urls += list(rerule.findall(res.text))

row = 0	#列（竖着）
col = 0 #行（横排）
file_obj = open('./book_info.txt', 'w', encoding='utf-8')
wb = xlwt.Workbook()
excel = wb.add_sheet('a test sheet')
# save book info
#书名 封面图 评分 作者 出版社 译者 出版年 页数 定价 装帧 ISBN 标签 简介 目录
for url in book_urls:
    # if(booknum <= 60):
       # continue
    if(booknum >= 200):
        break
    booknum = booknum+1
    if(booknum % 2 == 1):
        continue
    #file_obj = open('./book_info'+str(booknum/2)+'.txt', 'w', encoding='utf-8')
    try:
        res = requests.get(url, headers=headers)
        soup = BeautifulSoup(res.text, "html.parser")
        excel.write(col,row,soup.find(attrs={'property': 'v:itemreviewed'}).string+'\n')
    except:
        print("页面不存在")
        continue
    row += 1

    excel.write(col,row,soup.find(attrs={'class': 'nbg'}).get('href')+'\n')
    row += 1
    excel.write(col,row,soup.find(attrs={'class': 'll rating_num '}).text+'\n')
    row += 1
    temp = soup.find(attrs={'id': 'info'}).text
    temp = re.sub("\r","",temp);
    temp = re.sub(" ","",temp);
    #print(temp)
    num1 = temp.find("作者")
    while temp[num1] != '\n' or temp[num1+1] == '\n':
        #print(temp[num1])
        num1+=1
        continue
    num1+=1
    num2 = temp.index('\n',num1)
    excel.write(col,row,temp[num1:num2]+'\n')
    row += 1

    num1 = temp.find("出版社")
    num1 = temp.index(':',num1)+1
    num2 = temp.index('\n',num1)
    excel.write(col,row,temp[num1:num2]+'\n')
    row += 1
	
    num1 = temp.find("译者")
    if temp[num1] == '译':
        while temp[num1] != '\n' or temp[num1+1] == '\n':
            num1+=1
            continue
        num1+=1
        num2 = temp.index('\n',num1)
        excel.write(col,row,temp[num1:num2]+'\n')
    else:
        excel.write(col,row,'无'+'\n')
    row += 1
	
    num1 = temp.find("出版年")
    if temp[num1] == '出':
        num1 = temp.index(':',num1)+1
        num2 = temp.index('\n',num1)
        print(temp[num1:num2])
        excel.write(col,row,temp[num1:num2])
    else:
        excel.write(col,row,'无')
    row+=1
	
    num1 = temp.find("页数")
    if temp[num1] == '页':
        num1 = temp.index(':',num1)+1
        num2 = temp.index('\n',num1)
        excel.write(col,row,temp[num1:num2]+'\n')
    else:
        excel.write(col,row,'无'+'\n')
    row+=1

    num1 = temp.find("定价")
    if temp[num1] == '定':
        num1 = temp.index(':',num1)+1
        num2 = temp.index('\n',num1)
        excel.write(col,row,temp[num1:num2]+'\n')
    else:
        excel.write(col,row,'无'+'\n')
    row += 1
	
    num1 = temp.find("装帧")
    if temp[num1] == '装':
        num1 = temp.index(':',num1)+1
        num2 = temp.index('\n',num1)
        excel.write(col,row,temp[num1:num2]+'\n')
    else:
        excel.write(col,row,'无'+'\n')
    row+=1

    
    num1 = temp.find("ISBN")
    if temp[num1] == 'I':
        num1 = temp.index(':',num1)+1
        num2 = len(temp)-1
        excel.write(col,row,temp[num1:num2]+'\n')
    else:
        excel.write(col,row,'无' + '\n')
    row += 1
	
    temp = soup.find_all(attrs={'class': 'tag'})
    tags = ''
    for i in temp:
        # excel.write(col,row,i.text+' ')
        tags += i.text
    excel.write(col,row,tags)
    row += 1
    
    temp = soup.find_all(attrs={'class': 'intro'})
    # temp = soup.find_all("p","intro")
    # print(temp)
    if len(temp) >= 1: 
        t = temp[len(temp)-1].text
        # print(t)
        t2 = temp[0].text
        # print(t2)
        excel.write(col,row,t)
        row += 1
        excel.write(col,row,t2)
    else:
        excel.write(col,row,"无")
        row += 1
        excel.write(clo,row,"无")
    row += 1

    try:
        temp = soup.find(id = re.compile("dir_[0-9]*_full"))
        temp = temp.text
        temp = re.sub("\n","",temp);
        temp = re.sub("\r","",temp);
        temp = re.sub(" ","",temp);
        file_obj.write(temp+'\n')
    except:
        #print("目录错误")
        file_obj.write("本书暂无目录"+'\n')
    print(booknum/2)
    file_obj.write('\n')
    #file_obj.close()
    time.sleep(1)
    row = 0
    col += 1
file_obj.close()
wb.save('test.xls')
