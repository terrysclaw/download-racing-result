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
worksheet.write('C1', '班次')
worksheet.write('D1', '路程')
worksheet.write('E1', '草泥')

worksheet.write('F1', '馬號')
worksheet.write('G1', '馬名')
worksheet.write('H1', '騎師')
worksheet.write('I1', '練馬師')

worksheet.write('J1', '馬號')
worksheet.write('K1', '馬名')
worksheet.write('L1', '騎師')
worksheet.write('M1', '練馬師')

worksheet.write('N1', '馬號')
worksheet.write('O1', '馬名')
worksheet.write('P1', '騎師')
worksheet.write('Q1', '練馬師')

worksheet.write('R1', 'Win')
worksheet.write('S1', 'Qin')

sheet2 = workbook.add_worksheet()
sheet2.write('A1', 'Date')
sheet2.write('B1', '騎師王')
sheet2.write('C1', '練馬師王')



row = 1
row2 = 1

driver = webdriver.Chrome("chromedriver.exe")

for date in pd.date_range(start="2023-09-10", end="2024-02-28"):
    try:
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
            worksheet.write(row, 1, divs[0].text.replace('第', '').replace('場', ''))

            # race title
            title = divs[1].text
            worksheet.write(row, 2, title.split('-')[0])
            worksheet.write(row, 3, title.split('-')[1])
            if '草地' in title:
                worksheet.write(row, 4, '草地')
            elif '全天候跑道' in title:
                worksheet.write(row, 4, '全天候跑道')


            result = divs[3].find_all('div')
            trs = result[0].find_all('tr')
            
            tds = trs[1].find_all('td')
            worksheet.write(row, 5, int(tds[1].text))
            worksheet.write(row, 6, tds[2].text)
            worksheet.write(row, 7, tds[3].text)
            worksheet.write(row, 8, tds[4].text)

            tds = trs[2].find_all('td')
            worksheet.write(row, 9, int(tds[1].text))
            worksheet.write(row, 10, tds[2].text)
            worksheet.write(row, 11, tds[3].text)
            worksheet.write(row, 12, tds[4].text)

            tds = trs[3].find_all('td')
            worksheet.write(row, 13, int(tds[1].text))
            worksheet.write(row, 14, tds[2].text)
            worksheet.write(row, 15, tds[3].text)
            worksheet.write(row, 16, tds[4].text)


            trs = result[1].find_all('tr')

            for i in range(len(trs)):
                tds = trs[i].find_all('td')
                
                if(len(tds)==3):
                    if tds[0].text == '獨贏':
                        worksheet.write(row, 17, float(tds[2].text.replace(',', '')))

                    if tds[0].text == '連贏':
                        worksheet.write(row, 18, float(tds[2].text.replace(',', '')))

                    if '騎師王' in tds[0].text:
                        sheet2.write(row2, 0, d)
                        sheet2.write(row2, 1, tds[1].text)

                    if '練馬師王' in tds[0].text:
                        sheet2.write(row2, 2, tds[1].text)
                        row2 += 1
            
            row += 1

    except:
        continue



try:
    # 練馬師王賠率走勢圖
    sheet3 = workbook.add_worksheet()
    sheet3.write('A1', '賽事日期')
    sheet3.write('B1', '練馬師')
    sheet3.write('C1', '馬場')
    sheet3.write('D1', '賽日積分')
    sheet3.write('E1', '開盤賠率')

    driver.get("https://racing.hkjc.com/racing/chinese/racing-info/tnc-odds-chart.aspx?season=2023")
    sleep(5)

    content = driver.page_source
    soup = BeautifulSoup(content, features="html.parser")

    div = soup.find_all('div', attrs={'id': 'raceTable'})
    trs = div[0].find_all('tr')
    for i in range(len(trs)):
        if i == 0:
            continue

        tds = trs[i].find_all('td')
        sheet3.write(i, 0, tds[0].text)
        sheet3.write(i, 1, tds[1].text)
        sheet3.write(i, 2, tds[2].text)
        sheet3.write(i, 3, float(tds[3].text))
        sheet3.write(i, 4, float(tds[4].text))


    
    # 騎師王賠率走勢圖
    sheet4 = workbook.add_worksheet()
    sheet4.write('A1', '賽事日期')
    sheet4.write('B1', '騎師')
    sheet4.write('C1', '馬場')
    sheet4.write('D1', '賽日積分')
    sheet4.write('E1', '開盤賠率')

    driver.get("https://racing.hkjc.com/racing/chinese/racing-info/jkc-odds-chart.aspx?season=2023")
    sleep(5)

    content = driver.page_source
    soup = BeautifulSoup(content, features="html.parser")

    div = soup.find_all('div', attrs={'id': 'raceTable'})
    trs = div[0].find_all('tr')
    for i in range(len(trs)):
        if i == 0:
            continue

        tds = trs[i].find_all('td')
        sheet4.write(i, 0, tds[0].text)
        sheet4.write(i, 1, tds[1].text)
        sheet4.write(i, 2, tds[2].text)
        sheet4.write(i, 3, float(tds[3].text))
        sheet4.write(i, 4, float(tds[4].text))

except:
    pass


sleep(3)
driver.quit()

workbook.close()
