"""
Description: 到'台灣證券交易所' 透過selenium選取 不同日期 的半導體產業
    再用bs4把資料抓取下來
    儲存至csv檔
Date: 2022/04/28
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time                         # time.sleep()
from bs4 import BeautifulSoup
import csv

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get('https://www.twse.com.tw/zh/page/trading/exchange/MI_INDEX.html') # 台灣證券交易所

type=driver.find_element(by=By.XPATH, value="//*[@id='main-form']/div/div/form/select")  # 找到'分類項目'的路徑
select=Select(type)    # 選取'分類項目'的下拉式選單
select.select_by_visible_text('半導體業') # 選取'半導體業'

yy=driver.find_element(by=By.XPATH, value='//*[@id="d1"]/select[1]')  # 找到'民國'的路徑
select1=Select(yy)    # 選取'民國'的下拉式選單
select1.select_by_visible_text('民國 111 年') # 選取'民國111年'

mm=driver.find_element(by=By.XPATH, value='//*[@id="d1"]/select[2]')  # 找到'月份'的路徑
select2=Select(mm)    # 選取'月份'的下拉式選單
select2.select_by_visible_text('04月') # 選取'04月'

def dataset(date,day):
    dd=driver.find_element(by=By.XPATH, value='//*[@id="d1"]/select[3]')  # 找到'日期'的路徑
    select3=Select(dd)        # 選取'日期'的下拉式選單
    select3.select_by_visible_text('{}日 ({})'.format(date,day))  # 選取 'date日 (day)'

    driver.find_element(by=By.XPATH, value='//*[@id="main-form"]/div/div/form/a').click()  # 按下'查詢'
    time.sleep(1)             # 停 1 sec

    length=driver.find_element(by=By.XPATH, value='//*[@id="report-table1_length"]/label/select')  # 找到'每頁x筆'的路徑
    select5=Select(length)    # 選取'每頁x筆'的下拉式選單
    select5.select_by_visible_text('全部') # 選取'全部'
    time.sleep(1)             # 停 1 sec

    str1 = driver.page_source
    soup = BeautifulSoup(str1, "html.parser")
    tr_list=soup.find_all("tr",{"class":["odd", "even"]}) # 找到每筆證券

    # 儲存至csv檔,以日期區分
    title=soup.find_all("th",class_="sorting_disabled")   # 找到title
    header=[]
    for th in title:
        header.append(th.string)
    rows=[header] # [[],[],[]]
    c = 0
    for tr in tr_list:
        td_list=tr.find_all("td",class_="dt-head-center")
        c += 1
        print("============================2022年4月{}日".format(date)+"第" + str(c) + "筆============================")
        row=[] # []
        for td in td_list:
            row.append(td.string)
        rows.append(row)
        print(row)      # 印出每筆證券資料
    with open('semiconductor_stock_dataset_111_04_{}.csv'.format(date), 'w', encoding='cp950') as f:
        writer = csv.writer(f,lineterminator='\n')
        for row in rows:
            writer.writerow(row)

dataset("25","一")
dataset("26","二")
dataset("27","三")
dataset("28","四")
driver.close()              # 視窗關閉