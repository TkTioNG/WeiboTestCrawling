# WeiboTestCrawling
By TKT for FYP testing purpose

# Prerequisites Installation
1. xlrd
2. xlwt
3. xlutis
4. selenium
```
$ pip install xlwt
$ pip install xlrd
$ pip install xlutis
$ pip install selenium
```

# Procedure
1. Insert Username and Password for Sina Weibo account in line 177-178
```
    username = "tankt-wp17@student.tarc.edu.my" #你的微博登录名
    password = "#########" #你的密码
```

2. Set file path for the crawled data output in line 180
```
    book_name_xls = "weibo_test.xls" #填写你想存放excel的路径，没有文件会自动创建
```
Note: Output file will be in .xls Excel file.
If file does not exist, a new .xls file will be produce.
If file does exist, output data will be written into the sheet which its name is same as current keyword search, if not, a new sheet will be produced.

3. Set Keywords for search results in line 183
```
    keywords = ["can", "high", "worry", "case"] #输入你想要的关键字，建议有超话的话加上##，如果结果较少，不加#
```
Keywords can be a list of keywords.

4. Run testCrawlMyself.py
