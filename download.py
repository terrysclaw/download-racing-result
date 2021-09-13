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
worksheet.write('J1', '1-3P')
worksheet.write('K1', '1-3Q')
worksheet.write('L1', '1-5P')
worksheet.write('M1', '1-5Q')
worksheet.write('N1', '潘莫P')
worksheet.write('O1', '潘莫田P')
worksheet.write('P1', '潘莫Q')
worksheet.write('Q1', '潘莫田Q')


row = 1

# driver = webdriver.Chrome("C:\\Users\\terry\\todo\\download-racing-result\\chromedriver.exe")
driver = webdriver.Chrome("chromedriver.exe")

for date in pd.date_range(start="2021-09-05",end="2021-09-12"):
    try:
        d = date.strftime("%Y/%m/%d")
        print(d)
        driver.get("https://racing.hkjc.com/racing/information/Chinese/Racing/ResultsAll.aspx?RaceDate=" + d)

        sleep(3)

        content = driver.page_source
        soup = BeautifulSoup(content, features="html.parser")

        races = soup.find_all('div', attrs={'class': 'f_fs13 margin_top15'})

        for race in races:
            J=K=L=M=N=O=P=Q=0

            divs = race.find_all('div')

            worksheet.write(row, 0, d)
            worksheet.write(row, 1, divs[0].text)
            worksheet.write(row, 2, divs[1].text)

            result = divs[3].find_all('div')
            trs = result[0].find_all('tr')
            
            tds = trs[1].find_all('td')
            worksheet.write(row, 3, int(tds[1].text))
            worksheet.write(row, 6, tds[3].text)
            if(int(tds[1].text)<4):
                J+=1
                K+=1
            if(int(tds[1].text)<6):
                L+=1
                M+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                N+=1
                O+=1
                P+=1
                Q+=1
            if(tds[3].text=='田泰安'):
                O+=1
                Q+=1

            tds = trs[2].find_all('td')
            worksheet.write(row, 4, int(tds[1].text))
            worksheet.write(row, 7, tds[3].text)
            if(int(tds[1].text)<4):
                J+=1
                K+=1
            if(int(tds[1].text)<6):
                L+=1
                M+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                N+=1
                O+=1
                P+=1
                Q+=1
            if(tds[3].text=='田泰安'):
                O+=1
                Q+=1

            tds = trs[3].find_all('td')
            worksheet.write(row, 5, int(tds[1].text))
            worksheet.write(row, 8, tds[3].text)
            if(int(tds[1].text)<4):
                J+=1
            if(int(tds[1].text)<6):
                L+=1
            if(tds[3].text=='潘頓' or tds[3].text=='莫雷拉'):
                N+=1
                O+=1
            if(tds[3].text=='田泰安'):
                O+=1


            worksheet.write(row, 9, J)
            worksheet.write(row, 10, K)
            worksheet.write(row, 11, L)
            worksheet.write(row, 12, M)
            worksheet.write(row, 13, N)
            worksheet.write(row, 14, O)
            worksheet.write(row, 15, P)
            worksheet.write(row, 16, Q)

            
            row += 1

    except:
        continue


sleep(3)
driver.quit()

workbook.close()
