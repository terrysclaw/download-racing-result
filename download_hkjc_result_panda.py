import requests
from bs4 import BeautifulSoup
# import xlsxwriter
import pandas as pd
from datetime import date, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time


# 设置 ChromeDriver 路径
chrome_driver_path = "chromedriver.exe"  # 替换为您的 ChromeDriver 路径

# 初始化 WebDriver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)


year = 2024

df = pd.DataFrame({})


## create date range from 2023/09/01 to 2024/07/14
dates_2023 = [
    {'date': date(2023, 9, 10), 'course': 'ST'},
    {'date': date(2023, 9, 13), 'course': 'HV'},
    {'date': date(2023, 9, 17), 'course': 'ST'},
    {'date': date(2023, 9, 20), 'course': 'HV'},
    {'date': date(2023, 9, 24), 'course': 'ST'},
    {'date': date(2023, 9, 27), 'course': 'HV'},
    {'date': date(2023, 10, 1), 'course': 'ST'},
    {'date': date(2023, 10, 4), 'course': 'HV'},
    {'date': date(2023, 10, 11), 'course': 'HV'},
    {'date': date(2023, 10, 15), 'course': 'ST'},
    {'date': date(2023, 10, 18), 'course': 'HV'},
    {'date': date(2023, 10, 22), 'course': 'ST'},
    {'date': date(2023, 10, 25), 'course': 'ST'},
    {'date': date(2023, 10, 29), 'course': 'HV'},
    {'date': date(2023, 11, 1), 'course': 'HV'},
    {'date': date(2023, 11, 5), 'course': 'ST'},
    {'date': date(2023, 11, 8), 'course': 'HV'},
    {'date': date(2023, 11, 11), 'course': 'ST'},
    {'date': date(2023, 11, 15), 'course': 'HV'},
    {'date': date(2023, 11, 19), 'course': 'ST'},
    {'date': date(2023, 11, 22), 'course': 'HV'},
    {'date': date(2023, 11, 26), 'course': 'ST'},
    {'date': date(2023, 11, 29), 'course': 'HV'},
    {'date': date(2023, 12, 3), 'course': 'ST'},
    {'date': date(2023, 12, 6), 'course': 'HV'},
    {'date': date(2023, 12, 10), 'course': 'ST'},
    {'date': date(2023, 12, 13), 'course': 'HV'},
    {'date': date(2023, 12, 17), 'course': 'ST'},
    {'date': date(2023, 12, 20), 'course': 'HV'},
    {'date': date(2023, 12, 23), 'course': 'ST'},
    {'date': date(2023, 12, 26), 'course': 'ST'},
    {'date': date(2023, 12, 29), 'course': 'HV'},
    {'date': date(2024, 1, 1), 'course': 'ST'},
    {'date': date(2024, 1, 4), 'course': 'HV'},
    {'date': date(2024, 1, 7), 'course': 'ST'},
    {'date': date(2024, 1, 10), 'course': 'HV'},
    {'date': date(2024, 1, 13), 'course': 'ST'},
    {'date': date(2024, 1, 17), 'course': 'HV'},
    {'date': date(2024, 1, 21), 'course': 'ST'},
    {'date': date(2024, 1, 24), 'course': 'ST'},
    {'date': date(2024, 1, 28), 'course': 'ST'},
    {'date': date(2024, 1, 31), 'course': 'HV'},
    {'date': date(2024, 2, 4), 'course': 'ST'},
    {'date': date(2024, 2, 7), 'course': 'HV'},
    {'date': date(2024, 2, 12), 'course': 'ST'},
    {'date': date(2024, 2, 15), 'course': 'HV'},
    {'date': date(2024, 2, 18), 'course': 'ST'},
    {'date': date(2024, 2, 21), 'course': 'HV'},
    {'date': date(2024, 2, 25), 'course': 'ST'},
    {'date': date(2024, 2, 28), 'course': 'HV'},
    {'date': date(2024, 3, 3), 'course': 'ST'},
    {'date': date(2024, 3, 6), 'course': 'HV'},
    {'date': date(2024, 3, 10), 'course': 'ST'},
    {'date': date(2024, 3, 13), 'course': 'HV'},
    {'date': date(2024, 3, 16), 'course': 'ST'},
    {'date': date(2024, 3, 20), 'course': 'HV'},
    {'date': date(2024, 3, 24), 'course': 'ST'},
    {'date': date(2024, 3, 27), 'course': 'HV'},
    {'date': date(2024, 3, 31), 'course': 'ST'},
    {'date': date(2024, 4, 3), 'course': 'ST'},
    {'date': date(2024, 4, 7), 'course': 'ST'},
    {'date': date(2024, 4, 10), 'course': 'HV'},
    {'date': date(2024, 4, 14), 'course': 'ST'},
    {'date': date(2024, 4, 17), 'course': 'HV'},
    {'date': date(2024, 4, 20), 'course': 'ST'},
    {'date': date(2024, 4, 24), 'course': 'HV'},
    {'date': date(2024, 4, 28), 'course': 'ST'},
    {'date': date(2024, 5, 1), 'course': 'HV'},
    {'date': date(2024, 5, 5), 'course': 'ST'},
    {'date': date(2024, 5, 8), 'course': 'HV'},
    {'date': date(2024, 5, 11), 'course': 'ST'},
    {'date': date(2024, 5, 15), 'course': 'HV'},
    {'date': date(2024, 5, 19), 'course': 'ST'},
    {'date': date(2024, 5, 22), 'course': 'HV'},
    {'date': date(2024, 5, 26), 'course': 'ST'},
    {'date': date(2024, 5, 29), 'course': 'ST'},
    {'date': date(2024, 6, 2), 'course': 'ST'},
    {'date': date(2024, 6, 5), 'course': 'HV'},
    {'date': date(2024, 6, 8), 'course': 'ST'},
    {'date': date(2024, 6, 12), 'course': 'HV'},
    {'date': date(2024, 6, 15), 'course': 'ST'},
    {'date': date(2024, 6, 23), 'course': 'ST'},
    {'date': date(2024, 6, 26), 'course': 'HV'},
    {'date': date(2024, 7, 1), 'course': 'ST'},
    {'date': date(2024, 7, 4), 'course': 'HV'},
    {'date': date(2024, 7, 6), 'course': 'ST'},
    {'date': date(2024, 7, 10), 'course': 'HV'},
    {'date': date(2024, 7, 14), 'course': 'ST'},
]

dates_2024 = [
    {'date': date(2024, 9, 8), 'course': 'ST'},
    {'date': date(2024, 9, 11), 'course': 'HV'},
    {'date': date(2024, 9, 15), 'course': 'ST'},
    {'date': date(2024, 9, 18), 'course': 'HV'},
    {'date': date(2024, 9, 22), 'course': 'ST'},
    {'date': date(2024, 9, 25), 'course': 'HV'},
    {'date': date(2024, 9, 28), 'course': 'ST'},
    {'date': date(2024, 10, 1), 'course': 'ST'},
    {'date': date(2024, 10, 6), 'course': 'ST'},
    {'date': date(2024, 10, 9), 'course': 'HV'},
    {'date': date(2024, 10, 13), 'course': 'ST'},
    {'date': date(2024, 10, 16), 'course': 'HV'},
    {'date': date(2024, 10, 20), 'course': 'ST'},
    {'date': date(2024, 10, 23), 'course': 'ST'},
    {'date': date(2024, 10, 27), 'course': 'HV'},
    {'date': date(2024, 10, 30), 'course': 'HV'},
    {'date': date(2024, 11, 3), 'course': 'ST'},
    {'date': date(2024, 11, 6), 'course': 'HV'},
    {'date': date(2024, 11, 9), 'course': 'ST'},
    {'date': date(2024, 11, 13), 'course': 'HV'},
    {'date': date(2024, 11, 17), 'course': 'ST'},
    {'date': date(2024, 11, 20), 'course': 'HV'},
    {'date': date(2024, 11, 24), 'course': 'ST'},
    {'date': date(2024, 11, 27), 'course': 'HV'},
    {'date': date(2024, 12, 1), 'course': 'ST'},
    {'date': date(2024, 12, 4), 'course': 'HV'},
    {'date': date(2024, 12, 8), 'course': 'ST'},
    {'date': date(2024, 12, 11), 'course': 'HV'},
    {'date': date(2024, 12, 15), 'course': 'ST'},
    {'date': date(2024, 12, 18), 'course': 'ST'},
    {'date': date(2024, 12, 22), 'course': 'ST'},
    {'date': date(2024, 12, 26), 'course': 'HV'},
    {'date': date(2024, 12, 29), 'course': 'ST'},
    {'date': date(2025, 1, 1), 'course': 'ST'},
    {'date': date(2025, 1, 5), 'course': 'ST'},
    {'date': date(2025, 1, 8), 'course': 'HV'},
    {'date': date(2025, 1, 12), 'course': 'ST'},
    {'date': date(2025, 1, 15), 'course': 'HV'},
    {'date': date(2025, 1, 19), 'course': 'ST'},
    {'date': date(2025, 1, 22), 'course': 'HV'},
    {'date': date(2025, 1, 26), 'course': 'ST'},
    {'date': date(2025, 1, 31), 'course': 'ST'},
]


# Loop through each date
for date in dates_2023 if year == 2023 else dates_2024:
    # Convert date
    race_date = date['date'].strftime('%Y/%m/%d')

    race_course = date['course']

    # URL to scrape HKJC racing results
    for race_no in range(1, 12):
        data = {
            '日期': [],
            '總場次': [],
            '馬場': [],
            '場次': [],
            '班次': [],
            '路程': [],
            '賽事名稱': [],
            '泥草': [],
            '賽道': [],
            '賽事時間1': [],
            '賽事時間2': [],
            '賽事時間3': [],
            '賽事時間4': [],
            '賽事時間5': [],
            '賽事時間6': [],
            '名次': [],
            '馬號': [],
            '馬名': [],
            '布號': [],
            '騎師': [],
            '練馬師': [],
            '實際負磅': [],
            '排位體重': [],
            '檔位': [],
            '頭馬距離': [],
            '沿途走位': [],
            '完成時間': [],
            '獨贏賠率': [],
            '第 1 段': [],
            '第 2 段': [],
            '第 3 段': [],
            '第 4 段': [],
            '第 5 段': [],
            '第 6 段': [],
        }

        print(race_date, date['course'], race_no)

        # URL to scrape HKJC racing results
        url = f'https://racing.hkjc.com/racing/information/Chinese/Racing/LocalResults.aspx?RaceDate={race_date}&Racecourse={race_course}&RaceNo={race_no}' 

        # Make a GET request to the URL
        response = requests.get(url)

        try:

            # Parse the HTML using BeautifulSoup
            soup = BeautifulSoup(response.text, 'html.parser')


            race_div = soup.find('div', class_='race_tab')
            tr = race_div.find_all('tr')

            # Extract race_no, race_index
            race_index = tr[0].find_all('td')[0].text.split('(')[1].replace(')', '')

            # Extract class, distance, and rating
            race_class = tr[2].find_all('td')[0].text.split('-')[0]

            # Safely extract distance
            distance_parts = tr[2].find_all('td')[0].text.split('-')
            if len(distance_parts) > 1:
                distance = distance_parts[1].split('米')[0].strip()
            else:
                distance = ''

            # Safely extract rating
            rating_parts = distance_parts[1].split('(') if len(distance_parts) > 1 else []
            if len(rating_parts) > 1:
                rating = rating_parts[1].replace(')', '').strip()
            else:
                rating = ''

            condition = tr[2].find_all('td')[2].text


            # Extract class, distance, and rating
            cup = tr[3].find_all('td')[0].text

            band = tr[3].find_all('td')[2].text

            time_1 = tr[4].find_all('td')[2].text.replace('(', '').replace(')', '')
            time_2 = tr[4].find_all('td')[3].text.replace('(', '').replace(')', '')
            time_3 = tr[4].find_all('td')[4].text.replace('(', '').replace(')', '')
            if len(tr[4].find_all('td'))>5:
                time_4 = tr[4].find_all('td')[5].text.replace('(', '').replace(')', '')
            else:
                time_4 = None
            if len(tr[4].find_all('td'))>6:
                time_5 = tr[4].find_all('td')[6].text.replace('(', '').replace(')', '')
            else:
                time_5 = None
            if len(tr[4].find_all('td'))>7:
                time_6 = tr[4].find_all('td')[7].text.replace('(', '').replace(')', '')
            else:
                time_6 = None


            split_time_1 = tr[5].find_all('td')[2].text.replace('(', '').replace(')', '')
            split_time_2 = tr[5].find_all('td')[3].text.replace('(', '').replace(')', '')
            split_time_3 = tr[5].find_all('td')[4].text.replace('(', '').replace(')', '')
            if len(tr[5].find_all('td'))>5:
                split_time_4 = tr[5].find_all('td')[5].text.replace('(', '').replace(')', '')
            else:
                split_time_4 = None
            if len(tr[5].find_all('td'))>6:
                split_time_5 = tr[5].find_all('td')[6].text.replace('(', '').replace(')', '')
            else:
                split_time_5 = None
            if len(tr[5].find_all('td'))>7:
                split_time_6 = tr[5].find_all('td')[7].text.replace('(', '').replace(')', '')
            else:
                split_time_6 = None
            # print(split_time_1, split_time_2, split_time_3, split_time_4, split_time_5, split_time_6)


            # Find the table containing the race results
            results_div = soup.find('div', class_='performance')


            # Loop through each row in the table
            for tr in results_div.find_all('tr')[1:]:
                
                # Get data from each cell
                cells = tr.find_all('td')

                data['日期'].append(race_date)
                data['總場次'].append(f"{year}{race_index.zfill(3)}")
                data['馬場'].append(race_course)
                data['場次'].append(int(race_no))
                data['班次'].append(race_class)
                data['路程'].append(int(distance))
                data['賽事名稱'].append(cup)
                if '草地' in band:
                    data['泥草'].append('草地')
                    data['賽道'].append(band.split(' - ')[1])
                else:
                    data['泥草'].append('泥地')
                    data['賽道'].append(None)
                data['賽事時間1'].append(time_1)
                data['賽事時間2'].append(time_2)
                data['賽事時間3'].append(time_3)
                data['賽事時間4'].append(time_4)
                data['賽事時間5'].append(time_5)
                data['賽事時間6'].append(time_6)
                
                position = cells[0].text.replace('\r', '').replace('\n', '')
                try:
                    data['名次'].append(int(position))
                except:
                    data['名次'].append(None)

                horse_no = cells[1].text.replace('\r', '').replace('\n', '')
                try:
                    data['馬號'].append(int(horse_no))
                except:
                    data['馬號'].append(None)
                
                data['馬名'].append(cells[2].text.replace('\xa0', '').split('(')[0])
                data['布號'].append(cells[2].text.replace('\xa0', '').split('(')[1].replace(')', ''))
                data['騎師'].append(cells[3].text.replace('\r', '').replace('\n', ''))
                data['練馬師'].append(cells[4].text.replace('\r', '').replace('\n', ''))

                actual_weight = cells[5].text.replace('\r', '').replace('\n', '')

                try:
                    data['實際負磅'].append(int(actual_weight))
                except:
                    data['實際負磅'].append(None)

                horse_weight = cells[6].text.replace('\r', '').replace('\n', '')                    
                try:
                    data['排位體重'].append(int(horse_weight))
                except:
                    data['排位體重'].append(None)

                draw_no = cells[7].text.replace('\r', '').replace('\n', '')
                try:
                    data['檔位'].append(int(draw_no))
                except:
                    data['檔位'].append(None)


                data['頭馬距離'].append(cells[8].text.replace('\r', '').replace('\n', ''))
                data['沿途走位'].append(cells[9].text.replace('\r', '').replace('\n', ''))
                data['完成時間'].append(cells[10].text.replace('\r', '').replace('\n', ''))

                odds = cells[11].text.replace('\r', '').replace('\n', '')
                try:
                    data['獨贏賠率'].append(float(odds))
                except:
                    data['獨贏賠率'].append(None)


                data['第 1 段'].append('')
                data['第 2 段'].append('')
                data['第 3 段'].append('')
                data['第 4 段'].append('')
                data['第 5 段'].append('')
                data['第 6 段'].append('')

        except:            
            pass


        # get sectional time
        url = f"https://racing.hkjc.com/racing/information/Chinese/Racing/DisplaySectionalTime.aspx?RaceDate={date['date'].strftime('%d/%m/%Y')}&RaceNo={race_no}"
        driver.get(url)

        # 等待页面加载完成（根据需要调整等待时间）
        time.sleep(3)  # 可以替换为更智能的等待方式，如 WebDriverWait

        # 获取页面内容
        page_source = driver.page_source

        # # 关闭浏览器
        # driver.quit()

        # 使用 BeautifulSoup 解析页面内容
        soup = BeautifulSoup(page_source, 'html.parser')

        # 查找目标表格
        race_table = soup.find('table', class_='table_bd')

        if race_table:
            # 提取表头
            headers = [th.text.strip() for th in race_table.find('thead').find_all('td')]

            # 提取表格数据
            rows = []
            for row in race_table.find('tbody').find_all('tr'):
                cells = [cell.text.strip() for cell in row.find_all('td')]
                rows.append(cells)

            for row in rows:
                code = row[2].split('(')[1].replace(')', '')
                
                # find the index of the horse in data['布號'] match the code
                for index, item in enumerate(data['布號']):
                    if item == code:
                        try:
                            data['第 1 段'][index] = row[3].split('\n')[1].replace('\n', '')
                        except:
                            pass
                        try:
                            data['第 2 段'][index] = row[4].split('\n')[1].replace('\n', '')
                        except:
                            pass
                        try:
                            data['第 3 段'][index] = row[5].split('\n')[1].replace('\n', '')
                        except:
                            pass
                        try:                        
                            data['第 4 段'][index] = row[6].split('\n')[1].replace('\n', '')
                        except:
                            pass
                        try:
                            data['第 5 段'][index] = row[7].split('\n')[1].replace('\n', '')
                        except:
                            pass
                        try:
                            data['第 6 段'][index] = row[8].split('\n')[1].replace('\n', '')
                        except:
                            pass

        else:
            print("未找到目标表格。")


        ## append data to df
        df = df.append(pd.DataFrame(data), ignore_index=True)


# 关闭浏览器
driver.quit()


# loop through each row
for index, row in df.iterrows():
    # get next match 馬名
    next_match = df[(df['馬名'] == row['馬名']) & (df['總場次'] > row['總場次'])].head(1)
    
    if not next_match.empty:
        df.loc[index, '再出總場次'] = next_match['總場次'].values[0]
        df.loc[index, '再出名次'] = next_match['名次'].values[0]
        df.loc[index, '再出騎師'] = next_match['騎師'].values[0]
        df.loc[index, '再出獨贏賠率'] = next_match['獨贏賠率'].values[0]
    else:
        df.loc[index, '再出總場次'] = None
        df.loc[index, '再出名次'] = None
        df.loc[index, '再出騎師'] = None
        df.loc[index, '再出獨贏賠率'] = None


# Write the DataFrame to an Excel file
df.to_excel(f'raceresult_{year}.xlsx', index=False)


