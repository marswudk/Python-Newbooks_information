import requests
from bs4 import BeautifulSoup
import time
import random
import openpyxl

headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.100 Safari/537.36'
}

# 分析網址：https://www.books.com.tw/web/books_nbtopm_01/?v=1&o=5
# nbtopm_01 = 文學小說，02=商業理財以此類推/ v=1 顯示模式(清單or圖文) / o=5為排序方式 = 暢銷度

# 先抓分類
home_url = 'https://www.books.com.tw/web/books_nbtopm_01/?v=1&o=5'
home_html = requests.get(home_url, headers=headers).text
home_soup = BeautifulSoup(home_html, 'html.parser')
home_res = home_soup.find('div', {'class': 'mod_b type02_l001-1 clearfix'})
kinds = len(home_res.select("a"))
# 拆分網址
url = 'https://www.books.com.tw/web/books_nbtopm_'
mode = '/?v=1&o=5'


# 抓文學類新書資訊，發現都在class = 'mod_a clearfix'中
books = home_soup.find('div', {'class': 'mod_a clearfix'})

#建立excel工作簿
workbook = openpyxl.Workbook()
#獲取第一個工作表
sheet = workbook.worksheets[0]

#對網址跑迴圈
for kind in range(1, kinds+1):
    kind_url = '%s%02d%s' % (url,kind,mode)
    kind_html = requests.get(kind_url,headers= headers).text
    kind_soup = BeautifulSoup(kind_html,'html.parser')
    kind_res = kind_soup.find('div', {'class': 'mod_a clearfix'})
    
    #若存在分頁
    if 'cnt_page' in kind_html:
        pages = int(kind_res.select('.page span')[0].text)

    #若不存在分頁    
    else:
        pages = 1
    for page in range(1,pages+1):
        #拆分url %s->字串 %02d->二位數，不足補0
        page_url = '%s%02d%s%s%s' % (url,kind,mode,'&page=',page)
        page_html = requests.get(page_url,headers = headers).text
        page_soup = BeautifulSoup(page_html,'html.parser')
        page_res = page_soup.find('div', {'class': 'mod_a clearfix'})
        page_items = page_res.select('.item')
        n = 0
        
        list_items = []
        for item in page_items:
            names = item.select('.msg a')[0].text  # 書名
            authors = item.select('.msg a')[1].text  # 作者
            publishes = item.select('.msg a')[2].text  # 出版社
            dates = item.select('.msg span')[0].text.split('：')[-1]  # 出版日期
            #replace(將空白字元以空字串取代)，同時strip()濾除前後空白
            contents = item.select('.txt_cont')[0].text.replace(" ","").strip()  # 簡述
            prices = item.select('.set2')[0].text  # 價格
            n+=1
            list_data = [names, authors, publishes, dates, contents, prices]
            list_title = ['書名','作者','出版社','出版日期','內容','價格']
            list_items.append(list_data)
            print('書名:',names)
            print('作者:',authors)
            print('出版社:',publishes)
            print('出版日期:',dates)
            print('內容:',contents)
            print(prices)
    #每完成一個分類，隨機休息5~10秒        
    time.sleep(random.randint(5,10))        
    sheet.append(list_title)
    for item in list_items:
        sheet.append(item)
    workbook.save('博客來新書排行.xlsx')
    