from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import requests
from bs4 import BeautifulSoup
import re
import xlwt
import xlrd
import xlutils
import pandas as pd
import datetime
import random 
# !pip install xlutils

def mySelenium_XC_AirTicket(depart_date_box, ends_city_box, file_path, cover=1, roll_time=3, new_row=1):
    
    '''
    file_path：原有数据地址
    cover：True为覆盖，False为追加
    roll_time：拖动页面的次数，其值越大，页面越可能展示完整的数据
    new_row：在excel table中写入数据的行坐标
    '''

    if cover == 0:
        data = xlrd.open_workbook(file_path)#读入表格
        ws = xlutils.copy.copy(data) #复制之前表里存在的数据
        table=ws.get_sheet(0)

    table = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = table.add_sheet('plane price',cell_overwrite_ok=True)
    col = ('index', 'test_date','depart_city','arrive_city','lead', 'month', 'day','week', 'airline_name','price', 'depart_airport','depart_time', 'arrive_airport','arrive_time')
    for i in range(0,len(col)):
        sheet.write(0,i,col[i])
    for ends_city in ends_city_box:
        for depart_date in depart_date_box:

            time.sleep(random.uniform(2,5))
            print('ends_city: ', ends_city, 'depart_date: ', depart_date)

            '''
            使用selenium爬取网页
            '''

            s_driver = Service("C:\\Users\\ThinkPad\\AppData\\Local\\Google\\Chrome\\Application\\chromedriver")
            driver = webdriver.Chrome(service=s_driver)
            driver.maximize_window()
            url = 'https://flights.ctrip.com/online/list/oneway-'+ ends_city +'?depdate='+ depart_date +'&cabin=y_s_c_f&adult=1&child=0&infant=0'
            driver.get(url)

            for i in range(roll_time):
                try:
                    a = driver.find_element_by_class_name("btn")
                    # a = driver.find_element(by=By.CLASS_NAME, value="btn") 
                    a.click()
                    print(111111)
                except:
                    try:
                        target = driver.find_elements_by_xpath("//*[contains(@class, 'airline-name')]")[-1]
                        driver.execute_script("arguments[0].scrollIntoView();", target)  # 拖动到可见的元素去
                        time.sleep(0.2)
                        print(2222222)
                    except:
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(0.2)
                        driver.execute_script("window.scrollTo(0,document.body.scrollTop=100)")
                        time.sleep(0.2)
                        print(3333333)
                

            elements_flight_part_0 = driver.find_elements(by=By.CLASS_NAME, value='flight-part')
            elements_flight_item_0 = driver.find_elements(by=By.CLASS_NAME, value='flight-item.domestic')
            elements_flight_part = elements_flight_part_0
            elements_flight_item = elements_flight_item_0

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


                print('try success, then rundriver.close()')
                driver.close()


                print('len1：', len(soup_flight_part))
                # print(soup_flight_part)
                print('len2：', len(html_flight_item))
                # print(html_flight_item)

                '''
                提取html元素
                '''

                test_date = str(datetime.date.today())
                print(str(test_date))

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
                    soup = soup_flight_item[i]
                    index[i] = i+1
                    price[i] = soup.findAll(name="span", attrs={"class" :"price"})[0].text[1:]
                    unit = soup.findAll(name="dfn")[0].text
                    price[i] = price[i] + unit
                    
                    print(index[i], price[i])



                airline_name = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    
                    airline_name[i] = soup.findAll(name="div", attrs={"class" : "airline-name" })[0].text
                    
                    print(airline_name[i])


                depart_airport = [0 for i in range(len(soup_flight_item))]
                depart_time = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    
                    depart = soup.findAll(name="div", attrs={"class" : "depart-box" })[0]
                    depart_airport[i] = depart.findAll(name="span", attrs={"id" : re.compile('departureFlightTrain') })[0].text 
                    depart_time[i] = depart.findAll(name="div", attrs={"class" : "time" })[0].text
                    
                    print(depart_airport[i], depart_time[i])


                arrive_airport = [0 for i in range(len(soup_flight_item))]
                arrive_time = [0 for i in range(len(soup_flight_item))]
                for i in range(len(soup_flight_item)):
                    soup = soup_flight_item[i]
                    
                    arrive = soup.findAll(name="div", attrs={"class" : "arrive-box" })[0]
                    arrive_airport[i] = arrive.findAll(name="span", attrs={"id" : re.compile('arrivalFlightTrain') })[0].text 
                    arrive_time[i] = arrive.findAll(name="div", attrs={"class" : "time" })[0].text
                    
                    print(arrive_airport[i], arrive_time[i])

            except:
                print('try failure, then run driver.close()')
                driver.close()
                index = [1]
                test_date = str(datetime.date.today())
                depart_city = ends_city.split('-')[0]
                arrive_city = ends_city.split('-')[1]
                lead = 'lead'
                month = depart_date.split('-')[1]
                day = depart_date.split('-')[2]
                week = 'week'
                airline_name = ['airline']
                price = ['price']
                depart_airport = ['depart_airport']
                depart_time = ['depart_time']
                arrive_airport = ['arrive_airport']
                arrive_time = ['arrive_time']

            '''
            存储到excel中
            '''

            for i in range(0,len(index)):
                row = 0
                sheet.write(i+new_row, row, index[i]);row=row+1 
                sheet.write(i+new_row, row, test_date);row=row+1 
                sheet.write(i+new_row, row, depart_city);row=row+1 
                sheet.write(i+new_row, row, arrive_city);row=row+1 
                sheet.write(i+new_row, row, lead);row=row+1 
                sheet.write(i+new_row, row, month);row=row+1 
                sheet.write(i+new_row, row, day);row=row+1 
                sheet.write(i+new_row, row, week);row=row+1 
                sheet.write(i+new_row, row, airline_name[i]);row=row+1 
                sheet.write(i+new_row, row, price[i]);row=row+1 
                sheet.write(i+new_row, row, depart_airport[i]);row=row+1 
                sheet.write(i+new_row, row, depart_time[i]);row=row+1 
                sheet.write(i+new_row, row, arrive_airport[i]);row=row+1 
                sheet.write(i+new_row, row, arrive_time[i]);row=row+1 
            new_row = new_row + len(index)
            print(new_row)

    return table



def myGenerate_EndsCity(city_box):
    ends_city_box = []
    for city1 in city_box:
        for city2 in city_box:
            if city1 != city2:
                ends_city_box.append(city1 + '-' + city2)
    return ends_city_box


# depart_date_box = ['2022-06-01', '2022-06-02']
# city_box = ['sha', 'hkg', 'bjs']
# ends_city_box = myGenerate_EndsCity(city_box)
depart_date_box = ['2022-06-01']
ends_city_box = ['sha-hkg']
print(depart_date_box)
print(ends_city_box)

savepath = r'C:\\Users\\ThinkPad\\jupyter_notebook\\数据挖掘\\2022计算社会科学\\携程机票爬虫' + '\plane_price.xls'
table = mySelenium_XC_AirTicket(depart_date_box, ends_city_box, savepath, cover=1, roll_time=2, new_row=1)
table.save(savepath)  

pd.read_excel(savepath)