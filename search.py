import sys
import os
from selenium import webdriver
import time
from datetime import date
import xlwt

if __name__ == "__main__":
    # folderName = time.strftime("%Y-%m-%d_%H_%M_%S", time.localtime())
    # os.makedirs("./"+folderName+"/")             #创建文件夹，名称=日期+时间
    # 开始搜索
    browser = webdriver.Chrome()
    browser.get('http://www.collegeboard.org')
    browser.maximize_window()

    timeout = 60
    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_id('view15_username').send_keys('effyt07')
            browser.find_element_by_id('view15_password').send_keys('ILBB2ne1!')
            browser.find_element_by_xpath(".// *[ @ id = 'profile'] / div / div[5] / div / div[2] / div / div / div / div / div[1] / form / div[3] / \
                                      div[2] / button").submit()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_xpath(
                ".// *[ @ id = 'profile'] / div / div[5] / div / div / div[2] / div / div / div / div / div / ul / li[1] / a").click()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    # 选择第三个按钮
    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_id("finishMyRegistration3").click()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    # continue
    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_id("continue").click()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    # zip code
    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_id("zipCode").send_keys('91101')
            browser.find_element_by_id("searchByZipOrCountry").click()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    max_time = time.time() + timeout
    while time.time() < max_time:
        try:
            browser.find_element_by_id("showAvailableOnly").click()
        except:
            print("error")
        else:
            break
        time.sleep(0.2)

    xls = xlwt.Workbook()
    sht1 = xls.add_sheet('Sheet1')
    # 添加字段
    sht1.write(0, 0, 'Location')
    sht1.write(0, 1, 'Test center')
    table_list = []

    while browser.find_element_by_id("testCenterSearchResults_next").is_enabled():
        txt = browser.find_element_by_xpath(".//*[@id='testCenterSearchResults']/tbody").text
        if txt in table_list:
            break
        else:
            table_list.append(txt)
            browser.find_element_by_xpath("//*[@id='testCenterSearchResults_next']/a").click()
        time.sleep(0.2)

    print(table_list)
    big_count = 1
    for i in table_list:
        txt2 = i.split('\n')
        small_count = 0
        for j in txt2:
            if small_count == 0:
                sht1.write(big_count, small_count, j[:-22])
            else:
                sht1.write(big_count, small_count, j)
            small_count += 1
            if small_count == 2:
                small_count = 0
                big_count += 1

    str = time.strftime("%Y-%m-%d_%H_%M_%S", time.localtime())
    xls.save('./' + str + '.xls')
