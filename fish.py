import requests
import time
from lxml import etree
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import xlwt #写入文件
import xlrd #打开excel文件
from retrying import retry


file = xlwt.Workbook(encoding='utf-8', style_compression=0)
# 新建一个sheet
sheet = file.add_sheet('data')

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"}
option = webdriver.ChromeOptions()
option.add_argument('--headless')
driver =webdriver.Chrome(chrome_options=option)
driver.delete_all_cookies()
driver.get('http://yqdata.cnfm.com.cn/Stat/Loss.aspx')
global variety_len,style_len,area_len
area_num = 208
style_num = 1
variety_num = 1
month = driver.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_ddlMonthEnd"]/option')
month_len = len(month)
month_num = 3
year = driver.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_ddlYear"]/option')
year_len = len(year)
year_num = 4
h = 1
s = 0
time.sleep(0.3)
while year_num:
    # 时间
    r = 1
    while r:
            try:
                driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_ddlYear"]/option[{}]'.format(year_num)).click()
                time.sleep(0.3)
            except Exception:
                pass
            else:
               r=None
    try:
        r = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_ddlYear"]/option[{}]'.format(year_num)).text
    except Exception:
        print('年份出错')
        pass
    #月份
    while month_num:
        m = 1
        while m:
            try:
                driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_ddlMonthEnd"]/option[{}]'.format(month_num)).click()
                time.sleep(0.3)
            except Exception:
                print('月份出错')
                pass
            else:
               m=None
        try:
            mon = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_ddlMonthEnd"]/option[{}]'.format(month_num)).text
        except Exception:
            pass
        tim = r + '.'+mon
        # 所属地区
        while area_num:
            z = 1
            while z:
                try:
                    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbArea"]/option[{}]'.format(area_num)).click()
                    time.sleep(0.3)
                    area = driver.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_lsbArea"]/option')
                    area_len = len(area)
                    print(area_num)
                except Exception:
                    print('所属地区出错')
                    pass
                else:
                    z = None
            try:
                area = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbArea"]/option[{}]'.format(area_num)).text
            except Exception:
                pass
            # 养殖方式
            while style_num:
                y = 1
                while y:
                    try:
                        driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbStyle"]/option[{}]'.format(style_num)).click()
                        time.sleep(0.3)
                        style_len = len(driver.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_lsbStyle"]/option'))
                        print(style_num)
                    except Exception:
                        print('养殖方式出错')
                        pass
                    else:
                        y = None
                try:
                    idea = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbStyle"]/option[{}]'.format(style_num)).text
                except Exception:
                    print('养殖方式出错')
                    pass
                while variety_num:
                    # 养殖品种
                    x = 1
                    while x:
                        try:
                            driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbVariety"]/option[{}]'.format(variety_num)).click()
                            variety = driver.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_lsbVariety"]/option')
                            variety_len = len(variety)
                            print(variety_len)
                        except Exception:
                            pass
                        else:
                            x=None
                    try:
                        time.sleep(0.3)
                        zhong = driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_lsbVariety"]/option[{}]'.format(variety_num)).text
                    except Exception:
                        print('品种出错')
                        pass
                    # 查询
                    o = 1
                    while o:
                        try:
                            driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_btnSearch"]').click()
                            time.sleep(0.3)
                        except Exception:
                            print('点击出错')
                            pass
                        else:
                            o = None
                    try:
                        ll = driver.find_element_by_xpath(
                            '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]').text
                    except Exception:
                        pass
                    @retry(stop_max_attempt_number=2)
                    def dd():
                        global h,s
                        # 数据
                        # time.sleep(1)
                        if ll=='':
                            return
                        try:
                            g = driver.find_elements_by_xpath(
                                '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr')
                            g = len(g)
                        except Exception:
                            pass
                        while h:
                            try:
                                data1 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[1]'.format(
                                        h)).text
                                data2 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[2]'.format(
                                        h)).text
                                data3 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[3]'.format(
                                        h)).text
                                data4 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[4]'.format(
                                        h)).text
                                data5 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[5]'.format(
                                        h)).text
                                data6 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[6]'.format(
                                        h)).text
                                data7 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[7]'.format(
                                        h)).text
                                data8 = driver.find_element_by_xpath(
                                    '//*[@id="ContentPlaceHolder1_ctl00"]/div[2]/div/div/table/tbody[2]/tr[{}]/td[8]'.format(
                                        h)).text
                                print(tim+data1 + data2 + data3 + data4 + data5 + data6 + data7 + data8)
                                sheet.write(s, 0, tim)
                                sheet.write(s, 1, area)
                                sheet.write(s, 2, idea)
                                sheet.write(s, 3, zhong)
                                sheet.write(s, 4, data1)
                                sheet.write(s, 5, data2)
                                sheet.write(s, 6, data3)
                                sheet.write(s, 7, data4)
                                sheet.write(s, 8, data5)
                                sheet.write(s, 9, data6)
                                sheet.write(s, 10, data7)
                                sheet.write(s, 11, data8)
                                s += 1
                                if h == g:
                                    h = 1
                                    break
                                h += 1
                            except Exception:
                                print('完全出错')
                                return
                        file.save('yuye.xls')
                    time.sleep(0.3)
                    dd()
                            # return data1,data2,data3,data4,data5,data6,data7,data8
                    if ll == ''and variety_len < 91:
                        break
                    if ll==''and variety_num==1:
                        variety_num = 8
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==9:
                        variety_num = 40
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==41:
                        variety_num = 50
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==51:
                        variety_num = 59
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==60:
                        variety_num = 63
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==64:
                        variety_num = 77
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==78:
                        variety_num = 84
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    elif ll==''and variety_num==85:
                        variety_num = variety_len
                        if variety_num == variety_len:
                            variety_num = 1
                            print('种类完毕')
                            break
                    variety_num+=1
                    if variety_num == variety_len:
                        variety_num = 1
                        print('种类完毕')
                        break
                if style_num == style_len:
                    style_num = 1
                    print('养殖方式完毕')
                    break
                style_num += 1
            if area_num == area_len:
                area_num = 1
                print('地区执行完毕')
                break
            area_num+=1
        if month_num == month_len:
            month_num = 1
            print(str(month_num)+'执行完毕')
            break
        month_num +=1
    if year_num ==year_len:
        year_num = 1
        print(str(year_num)+'执行完毕')
        break
    year_num +=1
