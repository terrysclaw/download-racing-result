import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import numpy as np


# store the original df
# df_2023 = pd.read_excel('raceresult_2023.xlsx')
df_2024 = pd.read_excel('raceresult_2024.xlsx')
df_2025 = pd.read_excel('raceresult_2025.xlsx')

df_results = pd.concat([df_2024, df_2025], ignore_index=True)

df = df_results[(df_results['獨贏賠率'] <= 10) & (df_results['名次'] > 3) & (df_results['騎師'] == '潘頓')][['總場次', '馬名', '布號', '馬場', '泥草', '班次', '路程', '獨贏賠率', '騎師', '練馬師', '實際負磅', '名次', '完成時間', '獨贏賠率']]
# (df_results['名次'] <= 3) & 

for index, row in df.iterrows():
    last_match = df_results[(df_results['總場次'] > row['總場次']) & (df_results['布號'] == row['布號'])].head(1)


    if not last_match.empty and last_match['騎師'].values[0] == '潘頓':
        df.at[index, '上場總場次'] = last_match['總場次'].values[0]
        df.at[index, '上場名次'] = last_match['名次'].values[0]
        if row['路程'] >= 1000 and row['路程'] <= 1200:
            df.at[index, '上場賽事時間'] = last_match['賽事時間3'].values[0]
        elif row['路程'] > 1200 and row['路程'] <= 1650:
            df.at[index, '上場賽事時間'] = last_match['賽事時間4'].values[0]
        elif row['路程'] > 1650 and row['路程'] <= 2000:
            df.at[index, '上場賽事時間'] = last_match['賽事時間5'].values[0]
        elif row['路程'] > 2000:
            df.at[index, '上場賽事時間'] = last_match['賽事時間6'].values[0]
        df.at[index, '上場完成時間'] = last_match['完成時間'].values[0]

## df to excel
df.to_excel('last_match.xlsx', index=False)