from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import requests
from bs4 import BeautifulSoup
import re
import xlwt
import xlrd
import xlutils.copy
import pandas as pd
import datetime
import random 


def mySelenium_XC_AirTicket(depart_date_box, ends_city_box, file_path, cover=1, roll_time=3, new_row=1):
    
    '''
    file_path: 原有数据地址
    cover: True为覆盖，False为追加
    roll_time: 拖动页面的次数，其值越大，页面越可能展示完整的数据
    new_row: 在excel table中写入数据的行坐标
    '''

    if cover == 0:
        try: # 已存在同名表格时
            old_book = xlrd.open_workbook(file_path)#读入表格
            old_table = old_book.sheet_by_index(0)

            table = xlwt.Workbook(encoding='utf-8',style_compression=0) 
            sheet = table.add_sheet('plane price',cell_overwrite_ok=True)

            rows = old_table.nrows
            cols = old_table.ncols
            
            for i in range(0,rows):
                for j in range(0, cols):
                    print(i,j,old_table.cell_value(i, j))
                    sheet.write(i, j ,old_table.cell_value(i, j))
        except: # 未存在同名表格时

            table = xlwt.Workbook(encoding='utf-8',style_compression=0) 
            sheet = table.add_sheet('plane price',cell_overwrite_ok=True)
    else:
        table = xlwt.Workbook(encoding='utf-8',style_compression=0)
        sheet = table.add_sheet('plane price',cell_overwrite_ok=True)


    col = ('index', 'test_time','dep_city','arr_city','lead', 'month', 'day','week', 'airline_name', 'arrow_box', 'price', 'depart_airport','depart_time', 'arrive_airport','arrive_time')
    for i in range(0,len(col)):
        sheet.write(0,i,col[i])

    new_row = new_row + 1
    for ends_city in ends_city_box:
        for depart_date in depart_date_box:

            # time.sleep(random.uniform(2,5))
            print('ends_city: ', ends_city, 'depart_date: ', depart_date)

            '''
            使用selenium爬取网页
            '''
            options = webdriver.ChromeOptions() # 创建一个配置对象
            # options.add_argument("--headless") # 开启无界面模式
            # options.add_argument("--disable-gpu") # 禁用gpu
            # options.add_argument('--proxy-server=http://202.20.16.82:9527') # 使用代理ip
            # options.add_argument('--user-agent=Mozilla/5.0 HAHA') # 替换UA的命令 

            s = Service("C:\\Users\\ThinkPad\\AppData\\Local\\Google\\Chrome\\Application\\chromedriver")
            driver = webdriver.Chrome(service=s, options=options)
            driver.maximize_window()
            url = 'https://flights.ctrip.com/online/list/oneway-'+ ends_city +'?depdate='+ depart_date +'&cabin=y_s_c_f&adult=1&child=0&infant=0'
            print('url=', url)
            driver.get(url)

            for i in range(roll_time):

                alert_footer = driver.find_elements(by=By.CLASS_NAME, value="alert-footer")
                # print('alert_footer', alert_footer)

                if alert_footer != []: # 界面有警告框时
                    print('alert_footer[0].text', alert_footer[0].text)    
                    if alert_footer[0].text == '确认': # 第一次进入界面可能会展示一个“紧急公告”
                        alert_footer[0].click()    
                        print(111111)
                    elif alert_footer[0].text == '重新搜索': # 人工打码（三个验证框的文字都叫“重新搜索”）
                        while(alert_footer != []):
                            time.sleep(3)
                            print(222222)
                            # driver.execute_script("document.getElementById()")

                            alert_footer = driver.find_elements(by=By.CLASS_NAME, value="alert-footer")   
                
                else: # 界面没有警报框时
                    try:
                        target = driver.find_elements(by=By.XPATH, value="//*[contains(@class, 'airline-name')]")[-1]
                        driver.execute_script("arguments[0].scrollIntoView();", target)  # 拖动到可见的元素去
                        time.sleep(0.2)
                        print(333333)
                    except:
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(0.2)
                        driver.execute_script("window.scrollTo(0,document.body.scrollTop=100)")
                        time.sleep(0.2)
                        print(444444)
                

            elements_flight_part_0 = driver.find_elements(by=By.CLASS_NAME, value='flight-part')
            elements_flight_item_0 = driver.find_elements(by=By.CLASS_NAME, value='flight-item.domestic')
            elements_flight_part = elements_flight_part_0
            elements_flight_item = elements_flight_item_0


            # '''
            # 断点调试（不关闭浏览器）
            # '''
            # test = 1
            # while(test==1):
            #     print('testing')
            #     time.sleep(3)                
            #     a = driver.find_elements(by=By.CLASS_NAME, value="alert-footer")
            #     print(a) 
            #     # if a == []:
            #     #     test = 2


            '''
            解析html
            '''

            try:
                html_flight_part = elements_flight_part[0].get_attribute("outerHTML")
                soup_flight_part = BeautifulSoup(html_flight_part, 'html.parser')

                html_flight_item =  [0 for i in range(len(elements_flight_item))] 
                soup_flight_item =  [0 for i in range(len(elements_flight_item))] 
                for i in range(len(elements_flight_item)):
                    html_flight_item[i] = elements_flight_item[i].get_attribute("outerHTML")
                    soup_flight_item[i] = BeautifulSoup(html_flight_item[i], 'html.parser')


                print('try success, then run driver.close()')
                driver.close()


                print('len1：', len(soup_flight_part))
                # print(soup_flight_part)
                print('len2：', len(html_flight_item))
                # print(html_flight_item)

                '''
                提取html元素
                '''

                test_time = str(datetime.datetime.now())
                print(str(test_time))

                lead = soup_flight_part.findAll(name="div", attrs={"class" :"lead"})[0].text
                lead = (re.findall('[\u4e00-\u9fa5]+',lead)[0])
                depart_city = soup_flight_part.findAll(name="span", attrs={"class" :"depart"})[0].text
                arrive_city = soup_flight_part.findAll(name="span", attrs={"class" :"arrive"})[0].text
                date = soup_flight_part.find(name="div", attrs={"class" :"date"}).find(text=True).strip()
                month = date[:2]
                day = date[3:5]
                week = soup_flight_part.findAll(name="span", attrs={"class" :"week"})[0].text
                print(lead, depart_city, arrive_city, date, month, day, week)


                index = [0 for i in range(len(soup_flight_item))]
                price = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    try:
                        soup = soup_flight_item[i]
                        index[i] = i+1
                        price[i] = soup.findAll(name="span", attrs={"class" :"price"})[0].text[1:]
                        unit = soup.findAll(name="dfn")[0].text
                        price[i] = price[i] + unit
                    except:
                        index[i] ='none'
                        price[i] ='none'
                    
                    print(index[i], price[i])



                airline_name = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    
                    try:
                        airline_name[i] = soup.findAll(name="div", attrs={"class" : "airline-name" })[0].text
                    except:
                        airline_name[i] = 'none'         

                    print(airline_name[i])

                arrow_box = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    try:
                        arrow_box[i] = soup.findAll(name="div", attrs={"class" : "arrow-box" })[0].text
                        if not arrow_box[i]:
                            arrow_box[i] = '不中转'
                    except:
                        arrow_box[i] = 'none' 

                    print(arrow_box[i])


                depart_airport = [0 for i in range(len(soup_flight_item))]
                depart_time = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    try:
                        depart = soup.findAll(name="div", attrs={"class" : "depart-box" })[0]
                        depart_airport[i] = depart.findAll(name="span", attrs={"id" : re.compile('departureFlightTrain') })[0].text 
                        depart_time[i] = depart.findAll(name="div", attrs={"class" : "time" })[0].text
                    except:
                        depart_airport[i] = 'none' 
                        depart_time[i] = 'none' 
                    
                    print(depart_airport[i], depart_time[i])


                arrive_airport = [0 for i in range(len(soup_flight_item))]
                arrive_time = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    try:
                        arrive = soup.findAll(name="div", attrs={"class" : "arrive-box" })[0]
                        arrive_airport[i] = arrive.findAll(name="span", attrs={"id" : re.compile('arrivalFlightTrain') })[0].text 
                        arrive_time[i] = arrive.findAll(name="div", attrs={"class" : "time" })[0].text
                    except:
                        arrive_airport[i] = 'none' 
                        arrive_time[i] = 'none' 

                    print(arrive_airport[i], arrive_time[i])

            except:
                print('try failure, then run driver.close()')
                driver.close()

                index = [1]
                test_time = str(datetime.datetime.now())
                depart_city = ends_city.split('-')[0]
                arrive_city = ends_city.split('-')[1]
                lead = 'none'
                month = depart_date.split('-')[1]
                day = depart_date.split('-')[2]
                week = 'none'
                airline_name = ['none']
                arrow_box = ['none']
                price = ['none']
                depart_airport = ['none']
                depart_time = ['none']
                arrive_airport = ['none']
                arrive_time = ['none']

            '''
            存储到excel中
            '''

            for i in range(0,len(index)):
                column = 0
                sheet.write(i+new_row, column, index[i]);column=column+1 
                sheet.write(i+new_row, column, test_time);column=column+1 
                sheet.write(i+new_row, column, depart_city);column=column+1 
                sheet.write(i+new_row, column, arrive_city);column=column+1 
                sheet.write(i+new_row, column, lead);column=column+1 
                sheet.write(i+new_row, column, month);column=column+1 
                sheet.write(i+new_row, column, day);column=column+1 
                sheet.write(i+new_row, column, week);column=column+1 
                sheet.write(i+new_row, column, airline_name[i]);column=column+1 
                sheet.write(i+new_row, column, arrow_box[i]);column=column+1 
                sheet.write(i+new_row, column, price[i]);column=column+1 
                sheet.write(i+new_row, column, depart_airport[i]);column=column+1 
                sheet.write(i+new_row, column, depart_time[i]);column=column+1 
                sheet.write(i+new_row, column, arrive_airport[i]);column=column+1 
                sheet.write(i+new_row, column, arrive_time[i]);column=column+1 
            new_row = new_row + len(index)
            print('new_row=', new_row)


            table.save(file_path)  

    return table



def myGenerate_EndsCity_1(city_box):
    ends_city_box = []
    for city1 in city_box:
        for city2 in city_box:
            if city1 != city2:
                ends_city_box.append(city1 + '-' + city2)
    return ends_city_box


def myGenerate_EndsCity_2(dep_city):
    city_box = ['bjs', 'sha', 'can', 'szx', 'ctu', 'hgh', 
                'wuh', 'sia', 'ckg', 'tao', 'csx', 'nkg', 
                'xmn', 'kmg', 'dlc', 'tsn', 'cgo', 'syx' ]
    ends_city_box = []
    for arr_city in city_box:
        if dep_city != arr_city:
            ends_city_box.append(dep_city + '-' + arr_city)
    return ends_city_box



if __name__ == '__main__' :
    
    depart_date_box = ['2022-06-01']
    depart_city = 'hgh'
    ends_city_box = myGenerate_EndsCity_2(depart_city)

    print(depart_date_box)
    print(ends_city_box)

    savepath = r'C:\\Users\\ThinkPad\\jupyter_notebook\\数据挖掘\\2022计算社会科学\\携程机票爬虫\\机票价格_固定出发地版\\' + depart_city +'_depart_price.xlsx'
    print(savepath)

    table = mySelenium_XC_AirTicket(depart_date_box, ends_city_box, savepath, cover=0, roll_time=10, new_row=0)
    
    # table.save(savepath)  

    pd.read_excel(savepath)