import numpy as np
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from datetime import timedelta
from time import sleep
import xlsxwriter


workbook = xlsxwriter.Workbook('results.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Date')
worksheet.write('B1', 'Race No')
worksheet.write('C1', 'Race')


row = 1




driver = webdriver.Chrome("C:\\Users\\TerryLaw\\todo\\download-racing-result\\chromedriver.exe")

for date in pd.date_range(start="2020-09-01",end="2021-04-11"):
    try:

        # driver.get("https://racing.hkjc.com/racing/information/Chinese/Racing/ResultsAll.aspx?RaceDate=2021/04/11")
        d = date.strftime("%Y/%m/%d")
        print(d)
        driver.get("https://racing.hkjc.com/racing/information/Chinese/Racing/ResultsAll.aspx?RaceDate=" + d)

        sleep(3)

        content = driver.page_source
        soup = BeautifulSoup(content, features="html.parser")

        races = soup.find_all('div', attrs={'class': 'f_fs13 margin_top15'})

        for race in races:
            divs = race.find_all('div')

            worksheet.write(row, 0, d)
            worksheet.write(row, 1, divs[0].text)
            worksheet.write(row, 2, divs[1].text)

            result = divs[3].find_all('div')
            trs = result[0].find_all('tr')
            
            tds = trs[1].find_all('td')
            worksheet.write(row, 3, tds[1].text)
            worksheet.write(row, 4, tds[3].text)

            tds = trs[2].find_all('td')
            worksheet.write(row, 5, tds[1].text)
            worksheet.write(row, 6, tds[3].text)

            tds = trs[3].find_all('td')
            worksheet.write(row, 7, tds[1].text)
            worksheet.write(row, 8, tds[3].text)

            row += 1


    except:
        continue


sleep(3)
driver.quit()

workbook.close()
