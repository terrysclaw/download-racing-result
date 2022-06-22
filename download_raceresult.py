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
worksheet.write('M1', '1-3P')
worksheet.write('N1', '1-3Q')
worksheet.write('O1', '1-5P')
worksheet.write('P1', '1-5Q')
worksheet.write('Q1', '潘莫P')
worksheet.write('R1', '潘莫田P')
worksheet.write('S1', '潘莫Q')
worksheet.write('T1', '潘莫田Q')
worksheet.write('U1', 'Win')


row = 1

# driver = webdriver.Chrome("C:\\Users\\terry\\todo\\download-racing-result\\chromedriver.exe")
driver = webdriver.Chrome("chromedriver.exe")

for date in pd.date_range(start="2021-09-05",end="2022-06-19"):
    try:
        d = date.strftime("%Y/%m/%d")
        print(d)
        driver.get("https://racing.hkjc.com/racing/information/Chinese/Racing/ResultsAll.aspx?RaceDate=" + d)

        sleep(3)

        content = driver.page_source
        soup = BeautifulSoup(content, features="html.parser")

        races = soup.find_all('div', attrs={'class': 'f_fs13 margin_top15'})

        for race in races:
            M=N=O=P=Q=R=S=T=0

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
            if(int(tds[1].text)<4):
                M+=1
                N+=1
            if(int(tds[1].text)<6):
                O+=1
                P+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                Q+=1
                R+=1
                S+=1
                T+=1
            if(tds[3].text=='田泰安'):
                R+=1
                T+=1

            tds = trs[2].find_all('td')
            worksheet.write(row, 6, int(tds[1].text))
            worksheet.write(row, 7, tds[2].text)
            worksheet.write(row, 8, tds[3].text)
            if(int(tds[1].text)<4):
                M+=1
                N+=1
            if(int(tds[1].text)<6):
                O+=1
                P+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                Q+=1
                R+=1
                S+=1
                T+=1
            if(tds[3].text=='田泰安'):
                R+=1
                Q+=1

            tds = trs[3].find_all('td')
            worksheet.write(row, 9, int(tds[1].text))
            worksheet.write(row, 10, tds[2].text)
            worksheet.write(row, 11, tds[3].text)
            if(int(tds[1].text)<4):
                M+=1
            if(int(tds[1].text)<6):
                O+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                Q+=1
                R+=1
            if(tds[3].text=='田泰安'):
                R+=1


            worksheet.write(row, 12, M)
            worksheet.write(row, 13, N)
            worksheet.write(row, 14, O)
            worksheet.write(row, 15, P)
            worksheet.write(row, 16, Q)
            worksheet.write(row, 17, R)
            worksheet.write(row, 18, S)
            worksheet.write(row, 19, T)


            trs = result[1].find_all('tr')
            tds = trs[2].find_all('td')
            worksheet.write(row, 20, float(tds[2].text))

            
            row += 1

    except:
        continue


sleep(3)
driver.quit()

workbook.close()
