import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import numpy as np


# store the original df
df_2024 = pd.read_excel('raceresult_2024.xlsx')

df = df_2024[['總場次', '馬名', '布號', '馬場', '泥草', '路程', '獨贏賠率', '騎師', '練馬師', '實際負磅', '名次', '完成時間', '第 1 段', '第 2 段', '第 3 段', '第 4 段', '第 5 段', '第 6 段']]

for index, row in df.iterrows():
    if row['完成時間'] and row['完成時間'] != '---':
        time_parts = row['完成時間'].split(':')
        if len(time_parts) == 2:
            minutes, seconds = time_parts
            total_seconds = int(minutes) * 60 + float(seconds)
        else:
            total_seconds = float(row['完成時間'])

        df.at[index, '總時間'] = total_seconds

    if not np.isnan(row['第 6 段']):
        df.at[index, '最後800'] = row['第 6 段'] + row['第 5 段']
    elif not np.isnan(row['第 5 段']):
        df.at[index, '最後800'] = row['第 5 段'] + row['第 4 段']
    elif not np.isnan(row['第 4 段']):
        df.at[index, '最後800'] = row['第 4 段'] + row['第 3 段']
    elif not np.isnan(row['第 3 段']):
        df.at[index, '最後800'] = row['第 3 段'] + row['第 2 段']


for index, row in df.iterrows():
    ## found next match
    next_match = df[(df['總場次'] > row['總場次']) & (df['布號'] == row['布號']) & (df['馬場'] == row['馬場']) & (
            df['泥草'] == row['泥草']) & (df['獨贏賠率'] > 0)].head(1)
    
    if not next_match.empty:
        df.at[index, '下場總場次'] = next_match['總場次'].values[0]
        df.at[index, '下場馬場'] = next_match['馬場'].values
        df.at[index, '下場泥草'] = next_match['泥草'].values
        df.at[index, '下場路程'] = next_match['路程'].values
        df.at[index, '下場實際負磅'] = next_match['實際負磅'].values[0]
        df.at[index, '下場總時間'] = next_match['總時間'].values[0]
        df.at[index, '下場最後800'] = next_match['最後800'].values[0]
        df.at[index, '負磅調整'] = next_match['實際負磅'].values[0] - row['實際負磅']
        df.at[index, '總時間調整'] = next_match['總時間'].values[0] - row['總時間']
        df.at[index, '最後800調整'] = next_match['最後800'].values[0] - row['最後800']


## df to excel
df.to_excel('weight_adjusted_2024.xlsx', index=False)