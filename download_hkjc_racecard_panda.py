import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date


# store the original df
df_results = pd.read_excel('raceresult_2024.xlsx')

# 下載最新排位表
race_date = date(2025, 2, 5)
race_course = "HV"  # Happy Valley

# URL to scrape HKJC racing results
for race_no in range(1, 12):

    data = {
        '日期': [],
        '馬場': [],
        '場次': [],
        # '班次': [],
        # '路程': [],
        # '賽事名稱': [],
        # '泥草': [],
        # '賽道': [],
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


        race_div = soup.find('div', class_='f_fs13')
        

        # Find the table containing the race results
        racecard_table = soup.find('table', id='racecardlist')


        # Loop through each row in the table
        for tr in racecard_table.find_all('tbody')[0].find_all('tr')[2:]:
            
            # Get data from each cell
            cells = tr.find_all('td')

            data['日期'].append(race_date)
            data['馬場'].append(race_course)
            data['場次'].append(int(race_no))

            # data['班次'].append(race_class)
            # data['路程'].append(int(distance))
            # data['賽事名稱'].append(cup)
            # if '草地' in band:
            #     data['泥草'].append('草地')
            #     data['賽道'].append(band.split('草地')[1])
            # else:
            #     data['泥草'].append('泥地')
            #     data['賽道'].append(None)
            

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
        pass

    df = pd.DataFrame(data)

    ## loop through each row in df
    for index, row in df.iterrows():
        # get last match 馬名
        # last_match = df_results[(df_results['馬名'] == row['馬名']) & (df_results['總場次'] < row['總場次'])].tail(1)
        last_match = df_results[(df_results['布號'] == row['烙號'])].tail(1)

        if not last_match.empty:
            df.loc[index, '上次名次'] = last_match['名次'].values[0]
            df.loc[index, '上次騎師'] = last_match['騎師'].values[0]
            df.loc[index, '上次獨贏賠率'] = last_match['獨贏賠率'].values[0]
        else:
            df.loc[index, '上次名次'] = None
            df.loc[index, '上次騎師'] = None
            df.loc[index, '上次獨贏賠率'] = None


    ## export to excel file 
    df.to_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx", index=False)



output_file = f"Racecard_{race_date.strftime('%Y%m%d')}.xlsx"  # 替换为输出的 Excel 文件路径

# 创建一个 ExcelWriter 对象，用于写入多个工作表
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df = pd.read_excel("raceresult_2024.xlsx")

    df.to_excel(writer, sheet_name="result", index=False)


    for race_no in range(1, 12):
        df = pd.read_excel(f"Racecard_{race_date.strftime('%Y%m%d')}_{race_no}.xlsx")

        # 将 DataFrame 写入 Excel 文件中的新工作表
        df.to_excel(writer, sheet_name=f"Race_{race_no}", index=False)
