import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import numpy as np


# store the original df
df_2023 = pd.read_excel('raceresult_2023.xlsx')
df_2024 = pd.read_excel('raceresult_2024.xlsx')
df_2025 = pd.read_excel('raceresult_2025.xlsx')

df_results = pd.concat([df_2023, df_2024, df_2025], ignore_index=True)

# 去除班次欄位的前後空白，確保 '第四班 ' 能被修正為 '第四班'
df_results['班次'] = df_results['班次'].astype(str).str.strip()


# df = df_results[(df_results['騎師'] == '潘頓')][['總場次', '馬名', '布號', '馬場', '泥草', '班次', '路程', '獨贏賠率', '騎師', '練馬師', '實際負磅', '名次', '完成時間', '獨贏賠率']]
# df = df_results[(df_results['獨贏賠率'] < 10)][['總場次', '馬名', '布號', '馬場', '泥草', '班次', '路程', '騎師', '練馬師', '實際負磅', '名次', '完成時間', '獨贏賠率']]
df = df_results[(df_results['實際負磅'] > 129) & (df_results['名次'] < 4)][['總場次', '馬名', '布號', '馬場', '泥草', '班次', '路程', '騎師', '練馬師', '實際負磅', '名次', '完成時間', '獨贏賠率']]

for index, row in df.iterrows():
    # last_match = df_results[(df_results['總場次'] < row['總場次']) & (df_results['布號'] == row['布號'])].tail(1)
    # last_match = df_results[(df_results['總場次'] < row['總場次'])].tail(1)
    last_match = df_results[(df_results['總場次'] > row['總場次']) & (df_results['布號'] == row['布號'])].head(1)


    if not last_match.empty:
        df.at[index, '上次總場次'] = last_match['總場次'].values[0]
        df.at[index, '上次班次'] = last_match['班次'].values[0]
        df.at[index, '上次路程'] = last_match['路程'].values[0]
        df.at[index, '上次騎師'] = last_match['騎師'].values[0]
        df.at[index, '上次實際負磅'] = last_match['實際負磅'].values[0]
        df.at[index, '上次名次'] = last_match['名次'].values[0]
        df.at[index, '上次獨贏賠率'] = last_match['獨贏賠率'].values[0]

## remove if last match is empty
df = df.dropna(subset=['上次總場次'])

## df to excel
df.to_excel('last_match.xlsx', index=False)