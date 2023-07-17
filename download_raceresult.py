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
worksheet.write('M1', '1-4Q')
worksheet.write('N1', '5-9Q')
worksheet.write('O1', '10-14Q')
worksheet.write('P1', 'Win')
worksheet.write('Q1', 'Qin')

row = 1

driver = webdriver.Chrome("chromedriver.exe")

for date in pd.date_range(start="2023-05-14",end="2023-07-15"):
    try:
        d = date.strftime("%Y/%m/%d")
        print(d)
        driver.get("https://racing.hkjc.com/racing/information/Chinese/Racing/ResultsAll.aspx?RaceDate=" + d)

        sleep(3)

        content = driver.page_source
        soup = BeautifulSoup(content, features="html.parser")

        races = soup.find_all('div', attrs={'class': 'f_fs13 margin_top15'})

        for race in races:
            M=N=O=0

            divs = race.find_all('div')

            worksheet.write(row, 0, d)
            worksheet.write(row, 1, divs[0].text)
            worksheet.write(row, 2, divs[1].text)

            result = divs[3].find_all('div')
            trs = result[0].find_all('tr')
            
            tds = trs[1].find_all('td')
            worksheet.write(row, 3, int(tds[1].text))
            worksheet.write(row, 4, tds[2].text)
            worksheet.write(row, 5, tds[3].text)
            if(int(tds[1].text)<5):
                M+=1
            elif(int(tds[1].text)<10):
                N+=1
            elif(int(tds[1].text)>9):
                O+=1

            tds = trs[2].find_all('td')
            worksheet.write(row, 6, int(tds[1].text))
            worksheet.write(row, 7, tds[2].text)
            worksheet.write(row, 8, tds[3].text)
            if(int(tds[1].text)<5):
                M+=1
            elif(int(tds[1].text)<10):
                N+=1
            elif(int(tds[1].text)>9):
                O+=1

            tds = trs[3].find_all('td')
            worksheet.write(row, 9, int(tds[1].text))
            worksheet.write(row, 10, tds[2].text)
            worksheet.write(row, 11, tds[3].text)

            worksheet.write(row, 12, M)
            worksheet.write(row, 13, N)
            worksheet.write(row, 14, O)


            trs = result[1].find_all('tr')

            for i in range(len(trs)):
                tds = trs[i].find_all('td')
                if(len(tds)==3):
                    if tds[0].text == '獨贏':
                        worksheet.write(row, 15, float(tds[2].text.replace(',', '')))

                    if tds[0].text == '連贏':
                        worksheet.write(row, 16, float(tds[2].text.replace(',', '')))


            
            row += 1

    except:
        continue


sleep(3)
driver.quit()

workbook.close()
