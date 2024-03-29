# -*- coding: utf-8 -*-
"""
Created on Mon Nov 11 18:43:49 2019

@author: User
"""

import time
import xlrd
import xlwt
from xlutils.copy import copy
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# 定义一个滚动函数
def Transfer_Clicks(browser):
    try:
        browser.execute_script("window.scrollBy(0,document.body.scrollHeight)", "")
    except:
        pass
    return "Transfer successfully \n"

def isPresent():
    temp =1
    try: 
        elems = driver.find_elements_by_css_selector('div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
    except:
        temp =0
    return temp


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    #workbook = xlwt.Workbook()  # 新建一个工作簿
    try:
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        try:
            sheet = workbook.sheet_by_name(sheet_name)     
            new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
            new_worksheet = new_workbook.get_sheet(sheet_name)
        except :            
            new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
            new_worksheet = new_workbook.add_sheet(sheet_name)
        for i in range(0, index):
            for j in range(0, len(value[i])):
                new_worksheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
        new_workbook.save(path)
    except FileNotFoundError:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
        for i in range(0, index):
            for j in range(0, len(value[i])):
                sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
        workbook.save(path)
    print("xls格式表格写入数据成功！")
 
 
def write_excel_xls_append(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    #sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheet_name)  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(sheet_name)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")
 
 
def read_excel_xls(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
        print()
def spider(username,password,driver,book_name_xls,sheet_name_xls,keywords,maxWeibo):
         
    driver.set_window_size(1400, 800)
    driver.get("https://passport.weibo.cn/signin/login?entry=mweibo&res=wel&wm=3349&r=https%3A%2F%2Fm.weibo.cn%2F")
    #cookie1 = driver.get_cookies()
    #print (cookie1)
    time.sleep(2)
    elem = driver.find_element_by_xpath("//*[@id='loginName']");
    elem.send_keys(username)
    elem = driver.find_element_by_xpath("//*[@id='loginPassword']");
    elem.send_keys(password)
    elem = driver.find_element_by_xpath("//*[@id='loginAction']");
    elem.send_keys(Keys.ENTER)
    #cookie2 = driver.get_cookies()
    #print(cookie2)
    #获取信息
    while 1:  # 循环条件为1必定成立
        result = isPresent()
        print ('判断页面1成功 0失败  结果是=%d' % result )
        if result == 1:
            elems = driver.find_elements_by_css_selector('div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
            #return elems #如果封装函数，返回页面
            break
        else:
            print ('页面还没加载出来呢')
            time.sleep(20)
    time.sleep(2)
    elem = driver.find_element_by_xpath("//*[@class='m-text-cut']").click();
    time.sleep(2)
    elem = driver.find_element_by_xpath("//*[@type='search']");
    elem.send_keys(keywords)
    elem.send_keys(Keys.ENTER) 
    #handleScroll() 
    time.sleep(3)
    before = 0 
    after = 0
    n = 0 
    while True:
        before = after
        Transfer_Clicks(driver)
        time.sleep(2)
        elems = driver.find_elements_by_css_selector('div.card.m-panel.card9')
        print("当前包含微博最大数量：%d,n当前的值为：：%d, n值到30说明已无法解析出新的微博" % (len(elems),n))
        after = len(elems)
        if after > before:
            n = 0
        if after == before:        
            n = n + 1
        if n == 30:
            print("当前关键词最大微博数为：%d" % after)
            break
        if len(elems)>maxWeibo:
            break
        if len(elems) / 5000 == 0:
            time.sleep(300)
    
    value_title = [["rid", "用户名称", "微博等级", "微博内容", "微博转发量","微博评论量","微博点赞","发布时间","搜索关键词"],]
    write_excel_xls(book_name_xls, sheet_name_xls, value_title)
    rid = 0
    for elem in elems:
        rid = rid + 1
        #用户名
        weibo_username = elem.find_elements_by_css_selector('h3.m-text-cut')[0].text
        weibo_userlevel = "普通用户"
        #微博等级
        try: 
            weibo_userlevel_color_class = elem.find_elements_by_css_selector("i.m-icon")[0].get_attribute("class").replace("m-icon ","")
            if weibo_userlevel_color_class == "m-icon-yellowv":
                weibo_userlevel = "黄v"
            if weibo_userlevel_color_class == "m-icon-bluev":
                weibo_userlevel = "蓝v"
            if weibo_userlevel_color_class == "m-icon-goldv-static":
                weibo_userlevel = "金v"
            if weibo_userlevel_color_class == "m-icon-club":
                weibo_userlevel = "微博达人"     
        except:
            weibo_userlevel = "普通用户"
        #微博内容
        #分为有全文和无全文
        weibo_content = elem.find_elements_by_css_selector('div.weibo-text')[0].text
        shares = elem.find_elements_by_css_selector('i.m-font.m-font-forward + h4')[0].text
        comments = elem.find_elements_by_css_selector('i.m-font.m-font-comment + h4')[0].text
        likes = elem.find_elements_by_css_selector('i.m-icon.m-icon-like + h4')[0].text
        #发布时间
        weibo_time = elem.find_elements_by_css_selector('span.time')[0].text
#        print("用户名："+ weibo_username + "|"
#              "微博等级："+ weibo_userlevel + "|"
#              "微博内容："+ weibo_content + "|"
#              "转发："+ shares + "|"
#              "评论数："+ comments + "|"
#              "点赞数："+ likes + "|"
#              "发布时间："+ weibo_time + "|")
        value1 = [[rid, weibo_username, weibo_userlevel,weibo_content, shares,comments,likes,weibo_time,keywords],]
        write_excel_xls_append(book_name_xls, sheet_name_xls, value1)

if __name__ == '__main__':
    username = "tankt-wp17@student.tarc.edu.my" #你的微博登录名
    password = "#########" #你的密码
    driver = webdriver.Chrome('chromedriver.exe', service_log_path='NUL')#你的chromedriver的地址
    book_name_xls = "weibo_test.xls" #填写你想存放excel的路径，没有文件会自动创建
    #sheet_name_xls = '微博数据' #sheet表名
    maxWeibo = 10000 #设置最多多少条微博，如果未达到最大微博数量可以爬取当前已解析的微博数量
    keywords = ["can", "high", "worry", "case"] #输入你想要的关键字，建议有超话的话加上##，如果结果较少，不加#
    for i in range(len(keywords)):
        spider(username,password,driver,book_name_xls,keywords[i],keywords[i],maxWeibo)
        time.sleep(300)
