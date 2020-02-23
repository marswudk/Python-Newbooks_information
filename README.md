# Python-Newbooks_information_intoExcel
## 取得博客來新書排行榜資訊，並將其匯入Excel檔案中

專題目標：抓取博客來網站中，各分類新書排行榜的新書資訊，再將其匯入Excel檔。

1. 先解析網址，發現各分類的代碼、顯示模式及排序方式的規則。在此使用顯示模式v=1(清單)及o=5(依暢銷度排列)，所以差別在nbtopm後面的數字，為各分類代碼(01=文學小說、02=商業理財等)
```
# nbtopm_01 = 文學小說，02=商業理財以此類推/ v=1 顯示模式(清單or圖文) / o=5為排序方式 = 暢銷度
https://www.books.com.tw/web/books_nbtopm_01/?v=1&o=5

```

2. 先以nbtopm_01文學小說類的第一頁，找到所有類別，取得其長度，及所有新書的相關資訊位置
```
kinds = len(home_res.select("a")) #類別數
names = item.select('.msg a')[0].text  # 書名
authors = item.select('.msg a')[1].text  # 作者
publishes = item.select('.msg a')[2].text  # 出版社
dates = item.select('.msg span')[0].text.split('：')[-1]  # 出版日期
contents = item.select('.txt_cont p')[0]  # 簡述
prices = item.select('.set2')[0].text  # 價格
```

3. 拆分網址
```
url = 'https://www.books.com.tw/web/books_nbtopm_'
mode = '/?v=1&o=5'
```

4. 抓到第一類別的不同分頁

5. 對所有類別跑迴圈，以抓到不同分類的新書資訊，並加入判斷是否有分頁，沒有分頁就令page = 1

```
for kind in range(1, kinds+1):
    #%02d 將不足二位數的數字補0
    kind_url = '%s%02d%s' % (url,kind,mode)
```
```
if 'cnt_page' in kind_html:
        pages = int(kind_res.select('.page span')[0].text)

    #若不存在分頁    
    else:
        pages = 1
```

6. 以openpyxl模組將資料匯入Excel檔案中


