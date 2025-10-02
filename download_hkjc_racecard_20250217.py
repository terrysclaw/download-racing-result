import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import numpy as np


# store the original df
df_2024 = pd.read_excel('raceresult_2024.xlsx')
df_2025 = pd.read_excel('raceresult_2025.xlsx')

## concat the two df
df_results = pd.concat([df_2024, df_2025], ignore_index=True)


# 下載最新排位表
race_date = date(2025, 10, 4)
race_course = "ST"  # ST / HV

# URL to scrape HKJC racing results
for race_no in range(1, 12):

    data = {
        '日期': [],
        '馬場': [],
        '場次': [],
        '泥草': [],
        '賽道': [],
        '班次': [],
        '路程': [],
        '馬匹編號': [],
        '6次近績': [],
        '馬名': [],
        '烙號': [],
        '負磅': [],
        '騎師': [],
        '可能超磅': [],
        '檔位': [],
        '練馬師': [],
        '國際評分': [],
        '評分': [],
        '評分+/-': [],
        '排位體重': [],
        '排位體重+/-': [],
        '最佳時間': [],
        '馬齡': [],
        '分齡讓磅': [],
        '性別': [],
        '今季獎金': [],
        '優先參賽次序': [],
        '上賽距今日數': [],
        '配備': [],
        '馬主': [],
        '父系': [],
        '母系': [],
        '進口類別': [],
    }

    print(race_date, race_course, race_no)

    # URL to scrape HKJC racing results
    url = f"https://racing.hkjc.com/racing/information/Chinese/Racing/RaceCard.aspx?RaceDate={race_date.strftime('%Y/%m/%d')}&Racecourse={race_course}&RaceNo={race_no}"
    

    # Make a GET request to the URL
    response = requests.get(url)

    try:
        # Parse the HTML using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')

        # <div class="f_fs13" style="line-height: 20px;"><span class="font_wb" />第 1 場 - 鴻圖讓賽</span><br>2025年2月5日, 星期三, 跑馬地, 18:40<br>草地, "A" 賽道, 1650米, 好地<br>獎金: $875,000, 評分: 40-0, 第五班</div>

        race_div = soup.find('div', class_='f_fs13')

        if(race_div):

            band = None
            race_class = None
            distance = None

            for col in race_div.text.split(','):
                if '賽道' in col:
                    band = col
                if '班' in col:
                    race_class = col
                if '米' in col:
                    try:
                        distance = int(col.split('米')[0])
                    except:
                        distance = None
                        
            
            # Find the table containing the race results
            racecard_table = soup.find('table', id='racecardlist')

            # Loop through each row in the table
            for tr in racecard_table.find_all('tbody')[0].find_all('tr')[2:]:
                
                # Get data from each cell
                cells = tr.find_all('td')

                data['日期'].append(race_date)
                data['馬場'].append(race_course)
                data['場次'].append(int(race_no))
                data['泥草'].append('草地' if band else '泥地')
                data['賽道'].append(band)
                data['班次'].append(race_class)
                data['路程'].append(distance)

                try:
                    data['馬匹編號'].append(int(cells[0].text))
                except:
                    data['馬匹編號'].append(None)

                try:
                    data['6次近績'].append(cells[1].text)
                except:
                    data['6次近績'].append(None)

                try:
                    data['馬名'].append(cells[3].text)
                except:
                    data['馬名'].append(None)

                try:
                    data['烙號'].append(cells[4].text)
                except:
                    data['烙號'].append(None)

                try:
                    data['負磅'].append(int(cells[5].text))
                except:
                    data['負磅'].append(None)

                try:
                    data['騎師'].append(cells[6].text)
                except:
                    data['騎師'].append(None)

                try:
                    data['可能超磅'].append(cells[7].text)
                except:
                    data['可能超磅'].append(None)

                try:
                    data['檔位'].append(int(cells[8].text))
                except:
                    data['檔位'].append(None)

                try:
                    data['練馬師'].append(cells[9].text)
                except:
                    data['練馬師'].append(None)

                try:
                    data['國際評分'].append(int(cells[10].text))
                except:
                    data['國際評分'].append(None)

                try:
                    data['評分'].append(int(cells[11].text))
                except:
                    data['評分'].append(None)

                try:
                    data['評分+/-'].append(cells[12].text)
                except:
                    data['評分+/-'].append(None)

                try:
                    data['排位體重'].append(int(cells[13].text))
                except:
                    data['排位體重'].append(None)

                try:
                    data['排位體重+/-'].append(cells[14].text)
                except:
                    data['排位體重+/-'].append(None)

                try:
                    data['最佳時間'].append(cells[15].text)
                except:
                    data['最佳時間'].append(None)

                try:
                    data['馬齡'].append(int(cells[16].text))
                except:
                    data['馬齡'].append(None)

                try:
                    data['分齡讓磅'].append(int(cells[17].text))
                except:
                    data['分齡讓磅'].append(None)

                try:
                    data['性別'].append(cells[18].text)
                except:
                    data['性別'].append(None)

                try:
                    data['今季獎金'].append(cells[19].text)
                except:
                    data['今季獎金'].append(None)

                try:
                    data['優先參賽次序'].append(cells[20].text)
                except:
                    data['優先參賽次序'].append(None)

                try:
                    data['上賽距今日數'].append(cells[21].text)
                except:
                    data['上賽距今日數'].append(None)

                try:
                    data['配備'].append(cells[22].text)
                except:
                    data['配備'].append(None)

                try:
                    data['馬主'].append(cells[23].text)
                except:
                    data['馬主'].append(None)

                try:
                    data['父系'].append(cells[24].text)
                except:
                    data['父系'].append(None)

                try:
                    data['母系'].append(cells[25].text)
                except:
                    data['母系'].append(None)

                try:
                    data['進口類別'].append(cells[26].text)
                except:
                    data['進口類別'].append(None)
            
    except:            
        raise

    df = pd.DataFrame(data)


    ## loop through each row in df
    for index, row in df.iterrows():
        df.loc[index, '上次總場次'] = None
        df.loc[index, '上次日期'] = None
        df.loc[index, '上次馬場'] = None
        df.loc[index, '上次泥草'] = None
        df.loc[index, '上次場次'] = None
        df.loc[index, '上次班次'] = None
        df.loc[index, '上次路程'] = None
        df.loc[index, '上次場地狀況'] = None
        df.loc[index, '上次名次'] = None
        df.loc[index, '上次騎師'] = None
        df.loc[index, '上次賠率'] = None
        df.loc[index, '上次負磅'] = None
        df.loc[index, '上次負磅 +/-'] = None
        df.loc[index, '上次檔位'] = None
        df.loc[index, '上次檔位 +/-'] = None
        df.loc[index, '上次賽事時間1'] = None
        df.loc[index, '上次賽事時間2'] = None
        df.loc[index, '上次賽事時間3'] = None
        df.loc[index, '上次賽事時間4'] = None
        df.loc[index, '上次賽事時間5'] = None
        df.loc[index, '上次賽事時間6'] = None
        df.loc[index, '上次完成時間'] = None
        df.loc[index, '上次第 1 段'] = None
        df.loc[index, '上次第 2 段'] = None
        df.loc[index, '上次第 3 段'] = None
        df.loc[index, '上次第 4 段'] = None
        df.loc[index, '上次第 5 段'] = None
        df.loc[index, '上次第 6 段'] = None
        df.loc[index, '上次最後 800'] = None
        df.loc[index, '上次調整基數'] = None
        df.loc[index, '上次調整後最後 800'] = None
        df.loc[index, '上次調整後完成時間'] = None

        df.loc[index, '前次總場次'] = None
        df.loc[index, '前次日期'] = None
        df.loc[index, '前次馬場'] = None
        df.loc[index, '前次泥草'] = None
        df.loc[index, '前次場次'] = None
        df.loc[index, '前次班次'] = None
        df.loc[index, '前次路程'] = None
        df.loc[index, '前次場地狀況'] = None
        df.loc[index, '前次名次'] = None
        df.loc[index, '前次騎師'] = None
        df.loc[index, '前次賠率'] = None
        df.loc[index, '前次負磅'] = None
        df.loc[index, '前次負磅 +/-'] = None
        df.loc[index, '前次檔位'] = None
        df.loc[index, '前次檔位 +/-'] = None
        df.loc[index, '前次賽事時間1'] = None
        df.loc[index, '前次賽事時間2'] = None
        df.loc[index, '前次賽事時間3'] = None
        df.loc[index, '前次賽事時間4'] = None
        df.loc[index, '前次賽事時間5'] = None
        df.loc[index, '前次賽事時間6'] = None
        df.loc[index, '前次完成時間'] = None
        df.loc[index, '前次第 1 段'] = None
        df.loc[index, '前次第 2 段'] = None
        df.loc[index, '前次第 3 段'] = None
        df.loc[index, '前次第 4 段'] = None
        df.loc[index, '前次第 5 段'] = None
        df.loc[index, '前次第 6 段'] = None
        df.loc[index, '前次最後 800'] = None
        df.loc[index, '前次調整基數'] = None
        df.loc[index, '前次調整後最後 800'] = None
        df.loc[index, '前次調整後完成時間'] = None
        

        # 找上次同程紀錄
        last_match = df_results[(df_results['布號'] == row['烙號']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == row['路程']) & (df_results['獨贏賠率'] > 0)].tail(1)

        ## 如果沒有同程紀錄, 找相近路程紀錄 distance >= row['路程'] +- 250
        if last_match.empty:
            last_match = df_results[(df_results['馬名'] == row['馬名']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == int(row['路程']) - 200) & (df_results['獨贏賠率'] > 0)].tail(1)
        
        if last_match.empty:
            last_match = df_results[(df_results['馬名'] == row['馬名']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == int(row['路程']) + 200) & (df_results['獨贏賠率'] > 0)].tail(1)

        if last_match.empty:
            last_match = df_results[(df_results['馬名'] == row['馬名']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == int(row['路程']) - 150) & (df_results['獨贏賠率'] > 0)].tail(1)
        
        if last_match.empty:
            last_match = df_results[(df_results['馬名'] == row['馬名']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == int(row['路程']) + 150) & (df_results['獨贏賠率'] > 0)].tail(1)


        if not last_match.empty:
            df.loc[index, '上次總場次'] = last_match['總場次'].values[0]
            df.loc[index, '上次日期'] = last_match['日期'].values[0]
            df.loc[index, '上次馬場'] = last_match['馬場'].values[0]
            df.loc[index, '上次泥草'] = last_match['泥草'].values[0]
            df.loc[index, '上次場次'] = last_match['場次'].values[0]
            df.loc[index, '上次班次'] = last_match['班次'].values[0]
            df.loc[index, '上次路程'] = last_match['路程'].values[0]
            df.loc[index, '上次場地狀況'] = last_match['場地狀況'].values[0]
            df.loc[index, '上次名次'] = last_match['名次'].values[0]
            df.loc[index, '上次騎師'] = last_match['騎師'].values[0]
            df.loc[index, '上次賠率'] = last_match['獨贏賠率'].values[0]
            df.loc[index, '上次負磅'] = last_match['實際負磅'].values[0]
            ## calculate 
            if not np.isnan(last_match['實際負磅'].values[0]):
                df.loc[index, '上次負磅 +/-'] = row['負磅'] - last_match['實際負磅'].values[0]
            else:
                df.loc[index, '上次負磅 +/-'] = None

            df.loc[index, '上次檔位'] = last_match['檔位'].values[0]
            ## calculate 
            if not np.isnan(last_match['檔位'].values[0]):
                df.loc[index, '上次檔位 +/-'] = row['檔位'] - last_match['檔位'].values[0]
            else:
                df.loc[index, '上次檔位 +/-'] = None


            df.loc[index, '上次賽事時間1'] = last_match['賽事時間1'].values[0]
            df.loc[index, '上次賽事時間2'] = last_match['賽事時間2'].values[0]
            df.loc[index, '上次賽事時間3'] = last_match['賽事時間3'].values[0]
            df.loc[index, '上次賽事時間4'] = last_match['賽事時間4'].values[0]
            df.loc[index, '上次賽事時間5'] = last_match['賽事時間5'].values[0]
            df.loc[index, '上次賽事時間6'] = last_match['賽事時間6'].values[0]
            df.loc[index, '上次完成時間'] = last_match['完成時間'].values[0]
            df.loc[index, '上次第 1 段'] = last_match['第 1 段'].values[0]
            df.loc[index, '上次第 2 段'] = last_match['第 2 段'].values[0]
            df.loc[index, '上次第 3 段'] = last_match['第 3 段'].values[0]
            df.loc[index, '上次第 4 段'] = last_match['第 4 段'].values[0]
            df.loc[index, '上次第 5 段'] = last_match['第 5 段'].values[0]
            df.loc[index, '上次第 6 段'] = last_match['第 6 段'].values[0]
            df.loc[index, '上次最後 400'] = None
            df.loc[index, '上次最後 800'] = None

            ## calculate the last 800 and round to 2 decimal places
            if not np.isnan(last_match['第 6 段'].values[0]):
                df.loc[index, '上次最後 400'] = round(last_match['第 6 段'].values[0], 2)
                df.loc[index, '上次最後 800'] = round(last_match['第 5 段'].values[0] + last_match['第 6 段'].values[0], 2)
            elif not np.isnan(last_match['第 5 段'].values[0]):
                df.loc[index, '上次最後 400'] = round(last_match['第 5 段'].values[0], 2)
                df.loc[index, '上次最後 800'] = round(last_match['第 4 段'].values[0] + last_match['第 5 段'].values[0], 2)
            elif not np.isnan(last_match['第 4 段'].values[0]):
                df.loc[index, '上次最後 400'] = round(last_match['第 4 段'].values[0], 2)
                df.loc[index, '上次最後 800'] = round(last_match['第 3 段'].values[0] + last_match['第 4 段'].values[0], 2)
            elif not np.isnan(last_match['第 3 段'].values[0]):
                df.loc[index, '上次最後 400'] = round(last_match['第 3 段'].values[0], 2)
                df.loc[index, '上次最後 800'] = round(last_match['第 2 段'].values[0] + last_match['第 3 段'].values[0], 2)

            factor = 0
            if not np.isnan(df.loc[index, '上次負磅 +/-']):
                factor = df.loc[index, '上次負磅 +/-'] * 0.015

            if not np.isnan(df.loc[index, '上次檔位 +/-']):
                if df.loc[index, '馬場'] == 'HV':
                    if df.loc[index, '路程'] == 1000:
                        factor += df.loc[index, '上次檔位 +/-'] * 0.012
                    elif df.loc[index, '路程'] == 1200:
                        factor += df.loc[index, '上次檔位 +/-'] * 0.035
                    elif df.loc[index, '路程'] == 1650:
                        factor += df.loc[index, '上次檔位 +/-'] * 0.040
                    elif df.loc[index, '路程'] == 1800:
                        factor += df.loc[index, '上次檔位 +/-'] * 0.040
                    elif df.loc[index, '路程'] == 2200:
                        factor += df.loc[index, '上次檔位 +/-'] * 0.213
                else:
                    if df.loc[index, '泥草'] == '草地':
                        if df.loc[index, '路程'] == 1000:
                            factor += df.loc[index, '上次檔位 +/-'] * -0.019
                        elif df.loc[index, '路程'] == 1200:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.042
                        elif df.loc[index, '路程'] == 1400:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.018
                        elif df.loc[index, '路程'] == 1600:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.021
                        elif df.loc[index, '路程'] == 1800:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.056
                        elif df.loc[index, '路程'] == 2000:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.027
                        elif df.loc[index, '路程'] == 2400:
                            factor += df.loc[index, '上次檔位 +/-'] * -0.042
                    elif df.loc[index, '泥草'] == '泥地':
                        if df.loc[index, '路程'] == 1200:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.011
                        elif df.loc[index, '路程'] == 1650:
                            factor += df.loc[index, '上次檔位 +/-'] * 0.014
                        elif df.loc[index, '路程'] == 1800:
                            factor += df.loc[index, '上次檔位 +/-'] * -0.011

                
            df.loc[index, '上次調整基數'] = factor

            if df.loc[index, '上次最後 800'] is not None:
                df.loc[index, '上次調整後最後 800'] = df.loc[index, '上次最後 800'] + factor
            else:
                df.loc[index, '上次調整後最後 800'] = None

            ## convert 完成時間 to seconds
            if df.loc[index, '上次完成時間'] is None or df.loc[index, '上次完成時間'] == '---':
                df.loc[index, '上次調整後完成時間'] = None
                continue

            second = int(df.loc[index, '上次完成時間'].split(':')[0]) * 60 + float(df.loc[index, '上次完成時間'].split(':')[1])
            second += factor

            ## 如果上次路程與這次路程不同，計算整後完成時間
            if df.loc[index, '上次路程'] != df.loc[index, '路程']:
                if df.loc[index, '馬場'] == 'HV':
                    if df.loc[index, '路程'] == 1000 and df.loc[index, '上次路程'] == 1200: ## 1200 縮短至 1000
                        factor = 1 / 1.023
                    elif df.loc[index, '路程'] == 1200 and df.loc[index, '上次路程'] == 1000:   ## 1000 增程至 1200
                        factor = 1.023
                    elif df.loc[index, '路程'] == 1200 and df.loc[index, '上次路程'] == 1650:   ## 1650 縮短至 1200
                        factor = 1 / 1.041
                    elif df.loc[index, '路程'] == 1650 and df.loc[index, '上次路程'] == 1200:   ## 1200 增程至 1650
                        factor = 1.041
                    elif df.loc[index, '路程'] == 1650 and df.loc[index, '上次路程'] == 1800:   ## 1800 縮短至 1650
                        factor = 1 / 1.005
                    elif df.loc[index, '路程'] == 1800 and df.loc[index, '上次路程'] == 1650:   ## 1650 增程至 1800
                        factor = 1.005
                    elif df.loc[index, '路程'] == 1800 and df.loc[index, '上次路程'] == 2200:   ## 2200 縮短至 1800
                        factor = 1 / 1.023
                    elif df.loc[index, '路程'] == 2200 and df.loc[index, '上次路程'] == 1800:   ## 1800 增程至 2200
                        factor = 1.023

                else:
                    if df.loc[index, '泥草'] == '草地':
                        if df.loc[index, '路程'] == 1000 and df.loc[index, '上次路程'] == 1200: ## 1200 縮短至 1000
                            factor = 1 / 1.019
                        elif df.loc[index, '路程'] == 1200 and df.loc[index, '上次路程'] == 1000:   ## 1000 增程至 1200
                            factor = 1.019
                        elif df.loc[index, '路程'] == 1200 and df.loc[index, '上次路程'] == 1400:   ## 1400 縮短至 1200
                            factor = 1 / 1.014
                        elif df.loc[index, '路程'] == 1400 and df.loc[index, '上次路程'] == 1200:   ## 1200 增程至 1400
                            factor = 1.014
                        elif df.loc[index, '路程'] == 1400 and df.loc[index, '上次路程'] == 1600:   ## 1600 縮短至 1400
                            factor = 1 / 1.013
                        elif df.loc[index, '路程'] == 1600 and df.loc[index, '上次路程'] == 1400:   ## 1400 增程至 1600
                            factor = 1.013
                        elif df.loc[index, '路程'] == 1600 and df.loc[index, '上次路程'] == 1800:   ## 1800 縮短至 1600
                            factor = 1 / 1.011
                        elif df.loc[index, '路程'] == 1800 and df.loc[index, '上次路程'] == 1600:   ## 1600 增程至 1800
                            factor = 1.011
                        elif df.loc[index, '路程'] == 1800 and df.loc[index, '上次路程'] == 2000:   ## 2000 縮短至 1800
                            factor = 1 / 1.019
                        elif df.loc[index, '路程'] == 2000 and df.loc[index, '上次路程'] == 1800:   ## 1800 增程至 2000
                            factor = 1.019
                    else:
                        if df.loc[index, '路程'] == 1200 and df.loc[index, '上次路程'] == 1650: ## 1650 縮短至 1200
                            factor = 1 / 1.045
                        elif df.loc[index, '路程'] == 1650 and df.loc[index, '上次路程'] == 1200:   ## 1200 增程至 1650
                            factor = 1.045
                        elif df.loc[index, '路程'] == 1650 and df.loc[index, '上次路程'] == 1800:   ## 1800 縮短至 1650
                            factor = 1 / 1.007
                        elif df.loc[index, '路程'] == 1800 and df.loc[index, '上次路程'] == 1650:   ## 1650 增程至 1800
                            factor = 1.007

                second = second / df.loc[index, '上次路程'] * df.loc[index, '路程'] * factor


            ## convert 101.75 to 1:41.75
            df.loc[index, '上次調整後完成時間'] = f"{int(second / 60)}:{round(second % 60, 2)}"



            # 如果已經有上次同程紀錄，則找前次同程紀錄
            last_match = df_results[(df_results['布號'] == row['烙號']) & (df_results['馬場'] == row['馬場']) & (df_results['泥草'] == row['泥草']) & (df_results['路程'] == row['路程']) & (df_results['獨贏賠率'] > 0)].tail(2).head(1)

            if not last_match.empty and last_match['總場次'].values[0] != df.loc[index, '上次總場次']:
                df.loc[index, '前次總場次'] = last_match['總場次'].values[0]
                df.loc[index, '前次日期'] = last_match['日期'].values[0]
                df.loc[index, '前次馬場'] = last_match['馬場'].values[0]
                df.loc[index, '前次泥草'] = last_match['泥草'].values[0]
                df.loc[index, '前次場次'] = last_match['場次'].values[0]
                df.loc[index, '前次班次'] = last_match['班次'].values[0]
                df.loc[index, '前次路程'] = last_match['路程'].values[0]
                df.loc[index, '前次場地狀況'] = last_match['場地狀況'].values[0]
                df.loc[index, '前次名次'] = last_match['名次'].values[0]
                df.loc[index, '前次騎師'] = last_match['騎師'].values[0]
                df.loc[index, '前次賠率'] = last_match['獨贏賠率'].values[0]
                df.loc[index, '前次負磅'] = last_match['實際負磅'].values[0]

                ## calculate 
                if not np.isnan(last_match['實際負磅'].values[0]):
                    df.loc[index, '前次負磅 +/-'] = row['負磅'] - last_match['實際負磅'].values[0]
                else:
                    df.loc[index, '前次負磅 +/-'] = None

                df.loc[index, '前次檔位'] = last_match['檔位'].values[0]
                ## calculate 
                if not np.isnan(last_match['檔位'].values[0]):
                    df.loc[index, '前次檔位 +/-'] = row['檔位'] - last_match['檔位'].values[0]
                else:
                    df.loc[index, '前次檔位 +/-'] = None


                df.loc[index, '前次賽事時間1'] = last_match['賽事時間1'].values[0]
                df.loc[index, '前次賽事時間2'] = last_match['賽事時間2'].values[0]
                df.loc[index, '前次賽事時間3'] = last_match['賽事時間3'].values[0]
                df.loc[index, '前次賽事時間4'] = last_match['賽事時間4'].values[0]
                df.loc[index, '前次賽事時間5'] = last_match['賽事時間5'].values[0]
                df.loc[index, '前次賽事時間6'] = last_match['賽事時間6'].values[0]
                df.loc[index, '前次完成時間'] = last_match['完成時間'].values[0]
                df.loc[index, '前次第 1 段'] = last_match['第 1 段'].values[0]
                df.loc[index, '前次第 2 段'] = last_match['第 2 段'].values[0]
                df.loc[index, '前次第 3 段'] = last_match['第 3 段'].values[0]
                df.loc[index, '前次第 4 段'] = last_match['第 4 段'].values[0]
                df.loc[index, '前次第 5 段'] = last_match['第 5 段'].values[0]
                df.loc[index, '前次第 6 段'] = last_match['第 6 段'].values[0]
                df.loc[index, '前次最後 800'] = None

                ## calculate the last 800 and round to 2 decimal places
                if not np.isnan(last_match['第 6 段'].values[0]):
                    df.loc[index, '前次最後 400'] = round(last_match['第 6 段'].values[0], 2)
                    df.loc[index, '前次最後 800'] = round(last_match['第 5 段'].values[0] + last_match['第 6 段'].values[0], 2)
                elif not np.isnan(last_match['第 5 段'].values[0]):
                    df.loc[index, '前次最後 400'] = round(last_match['第 5 段'].values[0], 2)
                    df.loc[index, '前次最後 800'] = round(last_match['第 4 段'].values[0] + last_match['第 5 段'].values[0], 2)
                elif not np.isnan(last_match['第 4 段'].values[0]):
                    df.loc[index, '前次最後 400'] = round(last_match['第 4 段'].values[0], 2)
                    df.loc[index, '前次最後 800'] = round(last_match['第 3 段'].values[0] + last_match['第 4 段'].values[0], 2)
                elif not np.isnan(last_match['第 3 段'].values[0]):
                    df.loc[index, '前次最後 400'] = round(last_match['第 3 段'].values[0], 2)
                    df.loc[index, '前次最後 800'] = round(last_match['第 2 段'].values[0] + last_match['第 3 段'].values[0], 2)

                factor = 0
                if not np.isnan(df.loc[index, '前次負磅 +/-']):
                    factor = df.loc[index, '前次負磅 +/-'] * 0.015

                if not np.isnan(df.loc[index, '前次檔位 +/-']):
                    if df.loc[index, '馬場'] == 'HV':
                        if df.loc[index, '路程'] == 1000:
                            factor += df.loc[index, '前次檔位 +/-'] * 0.012
                        elif df.loc[index, '路程'] == 1200:
                            factor += df.loc[index, '前次檔位 +/-'] * 0.035
                        elif df.loc[index, '路程'] == 1650:
                            factor += df.loc[index, '前次檔位 +/-'] * 0.040
                        elif df.loc[index, '路程'] == 1800:
                            factor += df.loc[index, '前次檔位 +/-'] * 0.040
                        elif df.loc[index, '路程'] == 2200:
                            factor += df.loc[index, '前次檔位 +/-'] * 0.213
                    else:
                        if df.loc[index, '泥草'] == '草地':
                            if df.loc[index, '路程'] == 1000:
                                factor += df.loc[index, '前次檔位 +/-'] * -0.019
                            elif df.loc[index, '路程'] == 1200:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.042
                            elif df.loc[index, '路程'] == 1400:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.018
                            elif df.loc[index, '路程'] == 1600:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.021
                            elif df.loc[index, '路程'] == 1800:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.056
                            elif df.loc[index, '路程'] == 2000:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.027
                            elif df.loc[index, '路程'] == 2400:
                                factor += df.loc[index, '前次檔位 +/-'] * -0.042
                        elif df.loc[index, '泥草'] == '泥地':
                            if df.loc[index, '路程'] == 1200:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.011
                            elif df.loc[index, '路程'] == 1650:
                                factor += df.loc[index, '前次檔位 +/-'] * 0.014
                            elif df.loc[index, '路程'] == 1800:
                                factor += df.loc[index, '前次檔位 +/-'] * -0.011

                    
                df.loc[index, '前次調整基數'] = factor

                if df.loc[index, '前次最後 800'] is not None:
                    df.loc[index, '前次調整後最後 800'] = df.loc[index, '前次最後 800'] + factor
                else:
                    df.loc[index, '前次調整後最後 800'] = None

                ## convert 完成時間 to seconds
                if df.loc[index, '前次完成時間'] is None or df.loc[index, '前次完成時間'] == '---':
                    df.loc[index, '前次調整後完成時間'] = None
                    continue

                second = int(df.loc[index, '前次完成時間'].split(':')[0]) * 60 + float(df.loc[index, '前次完成時間'].split(':')[1])
                second += factor

                ## convert 101.75 to 1:41.75
                df.loc[index, '前次調整後完成時間'] = f"{int(second / 60)}:{round(second % 60, 2)}"


        ##############################################
        ## 計算2次調整後完成時間
        ##############################################
        df.loc[index, '馬名2'] = df.loc[index, '馬名']
        df.loc[index, '上次調整後完成時間2'] = df.loc[index, '上次調整後完成時間']
        if df.loc[index, '上次總場次'] is not None and not np.isnan(df.loc[index, '上次名次']) and not np.isnan(df.loc[index, '上次賠率']):
            df.loc[index, '上次賽事'] = f"{'##' if df.loc[index, '上次路程'] != df.loc[index, '路程'] else ''}[{pd.to_datetime(df.loc[index, '上次日期']).strftime('%d/%m')}] {int(df.loc[index, '上次路程'])} {df.loc[index, '上次騎師']} ({int(df.loc[index, '上次名次'])}) {df.loc[index, '上次賠率']}"
        else:
            df.loc[index, '上次賽事'] = None
        
        df.loc[index, '前次調整後完成時間2'] = df.loc[index, '前次調整後完成時間']
        if df.loc[index, '前次總場次'] is not None and not np.isnan(df.loc[index, '前次名次']) and not np.isnan(df.loc[index, '前次賠率']):
            df.loc[index, '前次賽事'] = f"[{pd.to_datetime(df.loc[index, '前次日期']).strftime('%d/%m')}] {int(df.loc[index, '前次路程'])} {df.loc[index, '前次騎師']} ({int(df.loc[index, '前次名次'])}) {df.loc[index, '前次賠率']}"
        else:
            df.loc[index, '前次賽事'] = None

        second = 0
        if df.loc[index, '上次總場次'] is not None:
            try:            
                second = int(df.loc[index, '上次調整後完成時間'].split(':')[0]) * 60 + float(df.loc[index, '上次調整後完成時間'].split(':')[1])

                try:
                    second += int(df.loc[index, '前次調整後完成時間'].split(':')[0]) * 60 + float(df.loc[index, '前次調整後完成時間'].split(':')[1])

                    second /= 2
                except:
                    pass
            except:
                pass


        if second > 0:
            df.loc[index, '2次調整後平均時間'] = f"{int(second / 60)}:{round(second % 60, 2)}"
        else:
            df.loc[index, '2次調整後平均時間'] = None

        ##
        second = 0
        if df.loc[index, '上次總場次'] is not None:
            try:
                second = int(df.loc[index, '上次調整後完成時間'].split(':')[0]) * 60 + float(df.loc[index, '上次調整後完成時間'].split(':')[1])
            except:
                pass

        if second > 0:
            df.loc[index, '上次調整後完成秒速'] = second
        else:
            df.loc[index, '上次調整後完成秒速'] = None

        
        second = 0
        if df.loc[index, '前次總場次'] is not None:
            try:
                second = int(df.loc[index, '前次調整後完成時間'].split(':')[0]) * 60 + float(df.loc[index, '前次調整後完成時間'].split(':')[1])
            except:
                pass

        if second > 0:
            df.loc[index, '前次調整後完成秒速'] = second
        else:
            df.loc[index, '前次調整後完成秒速'] = None


        second = 0
        if df.loc[index, '2次調整後平均時間'] is not None:
            try:
                second = int(df.loc[index, '2次調整後平均時間'].split(':')[0]) * 60 + float(df.loc[index, '2次調整後平均時間'].split(':')[1])
            except:
                pass

        if second > 0:
            df.loc[index, '2次調整後平均秒速'] = second
        else:
            df.loc[index, '2次調整後平均秒速'] = None

        ## calculate 2次較快完成秒速 min(上次調整後完成秒速, 前次調整後完成秒速)
        valid_times = [time for time in [df.loc[index, '上次調整後完成秒速'], df.loc[index, '前次調整後完成秒速']] if time is not None]
        if valid_times:
            if len(valid_times) == 1:
                df.loc[index, '近2次較快完成日期'] = df.loc[index, '上次日期']
                df.loc[index, '近2次較快完成馬場'] = df.loc[index, '上次馬場']
                df.loc[index, '近2次較快完成泥草'] = df.loc[index, '上次泥草']
                df.loc[index, '近2次較快完成場次'] = df.loc[index, '上次場次']
                df.loc[index, '近2次較快完成秒速'] = df.loc[index, '上次調整後完成秒速']
                df.loc[index, '近2次較快完成最後 400'] = df.loc[index, '上次最後 400']
                df.loc[index, '近2次較快完成最後 800'] = df.loc[index, '上次最後 800']
            else:
                df.loc[index, '近2次較快完成秒速'] = min(valid_times)
                if df.loc[index, '近2次較快完成秒速'] == valid_times[0]:
                    df.loc[index, '近2次較快完成日期'] = df.loc[index, '上次日期']
                    df.loc[index, '近2次較快完成馬場'] = df.loc[index, '上次馬場']
                    df.loc[index, '近2次較快完成泥草'] = df.loc[index, '上次泥草']
                    df.loc[index, '近2次較快完成場次'] = df.loc[index, '上次場次']
                    df.loc[index, '近2次較快完成最後 400'] = df.loc[index, '上次最後 400']
                    df.loc[index, '近2次較快完成最後 800'] = df.loc[index, '上次最後 800']
                else:
                    df.loc[index, '近2次較快完成日期'] = df.loc[index, '前次日期']
                    df.loc[index, '近2次較快完成馬場'] = df.loc[index, '前次馬場']
                    df.loc[index, '近2次較快完成泥草'] = df.loc[index, '前次泥草']
                    df.loc[index, '近2次較快完成場次'] = df.loc[index, '前次場次']
                    df.loc[index, '近2次較快完成最後 400'] = df.loc[index, '前次最後 400']
                    df.loc[index, '近2次較快完成最後 800'] = df.loc[index, '前次最後 800']
        else:
            df.loc[index, '近2次較快完成日期'] = None
            df.loc[index, '近2次較快完成馬場'] = None
            df.loc[index, '近2次較快完成泥草'] = None
            df.loc[index, '近2次較快完成場次'] = None
            df.loc[index, '近2次較快完成秒速'] = None
            df.loc[index, '近2次較快完成最後 400'] = None
            df.loc[index, '近2次較快完成最後 800'] = None


    ## export to excel file 
    df.to_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx", index=False)
    print(df.tail(10))



output_file = f"Racecard_{race_date.strftime('%Y%m%d')}.xlsx"  # 替换为输出的 Excel 文件路径

# 创建一个 ExcelWriter 对象，用于写入多个工作表
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for race_no in range(1, 12):
        df = pd.read_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx")

        # 将 DataFrame 写入 Excel 文件中的新工作表
        df.to_excel(writer, sheet_name=f"Race_{race_no}", index=False)



## Terry AI
output_file = f"Racecard_{race_date.strftime('%Y%m%d')}_Terry.xlsx"  # 替换为输出的 Excel 文件路径

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for race_no in range(1, 12):
        df = pd.read_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx")

        try:
            df_terry = df[['場次', '班次', '路程', '馬匹編號', '馬名', '騎師', '練馬師', '評分', '上次日期', '上次班次', '上次場地狀況', '負磅', '上次負磅', '上次負磅 +/-', '檔位', '上次檔位', '上次檔位 +/-', '最佳時間', '近2次較快完成秒速', '近2次較快完成最後 400', '近2次較快完成最後 800']]

            ## 在'上次日期'後面插入一列'上賽距今日數'
            df_terry.insert(df_terry.columns.get_loc('上次日期') + 1, '上賽距今日數', None)

            ## 在'上次完成時間'後面插入一列'彎位400'
            df_terry.insert(df_terry.columns.get_loc('近2次較快完成秒速') + 1, '彎位400', None)

            race_date_ts = pd.Timestamp(race_date)
            for index, row in df_terry.iterrows():
                if row['上次日期'] is not None and not pd.isna(row['上次日期']):
                    last_date = pd.to_datetime(row['上次日期'])
                    df_terry.at[index, '上賽距今日數'] = (race_date_ts - last_date).days

                ## 彎位400 = 上次最後 800 - 上次最後 400
                if row['近2次較快完成最後 800'] is not None and not pd.isna(row['近2次較快完成最後 800']) and row['近2次較快完成最後 400'] is not None and not pd.isna(row['近2次較快完成最後 400']):
                    df_terry.at[index, '彎位400'] = round(row['近2次較快完成最後 800'] - row['近2次較快完成最後 400'], 2)

                ## add a column '賽事重溫連結' with value
                df_terry.at[index, '賽事重溫連結'] = f"https://racing.hkjc.com/racing/video/play.asp?type=replay-full&date={race_date.strftime('%Y%m%d')}&no={race_no:02d}&lang=chi"
                # df_alfred.at[index, '賽事重溫連結'] = f"https://racing.hkjc.com/racing/video/play.asp?type=replay-full&date=20250420&no=06&lang=chi"

            print(df_terry.tail(10))

            df_terry.to_excel(writer, sheet_name=f"Race_{race_no}", index=False)
        except Exception as e:
            # print(f"Error processing Race sheet for {race_date} Race {race_no}: {e}")
            pass


## Alfred AI
output_file = f"Racecard_{race_date.strftime('%Y%m%d')}_Alfred.xlsx"  # 替换为输出的 Excel 文件路径

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for race_no in range(1, 12):
        df = pd.read_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx")

        try:
            df_alfred = df[['場次', '班次', '路程', '馬匹編號', '馬名', '騎師', '評分', '上次日期', '上次班次', '上次場地狀況', '負磅', '上次負磅', '上次負磅 +/-', '檔位', '上次檔位', '上次檔位 +/-', '最佳時間', '上次完成時間', '上次最後 400', '上次最後 800']]

            ## 在'上次日期'後面插入一列'上賽距今日數'
            df_alfred.insert(df_alfred.columns.get_loc('上次日期') + 1, '上賽距今日數', None)

            ## 在'上次完成時間'後面插入一列'灣位400'
            df_alfred.insert(df_alfred.columns.get_loc('上次完成時間') + 1, '灣位400', None)

            race_date_ts = pd.Timestamp(race_date)
            for index, row in df_alfred.iterrows():
                if row['上次日期'] is not None and not pd.isna(row['上次日期']):
                    last_date = pd.to_datetime(row['上次日期'])
                    df_alfred.at[index, '上賽距今日數'] = (race_date_ts - last_date).days

                ## 灣位400 = 上次最後 800 - 上次最後 400
                if row['上次最後 800'] is not None and not pd.isna(row['上次最後 800']) and row['上次最後 400'] is not None and not pd.isna(row['上次最後 400']):
                    df_alfred.at[index, '灣位400'] = round(row['上次最後 800'] - row['上次最後 400'], 2)

            print(df_alfred.tail(10))

            df_alfred.to_excel(writer, sheet_name=f"Alfred_{race_no}", index=False)
        except Exception as e:
            # print(f"Error processing Alfred sheet for {race_date} Race {race_no}: {e}")
            pass