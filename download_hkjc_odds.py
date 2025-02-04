import requests
import pandas as pd


output_file = f"HKJC_odds.xlsx"  # 替换为输出的 Excel 文件路径

# 创建一个 ExcelWriter 对象，用于写入多个工作表
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for race_no in range(1, 12):
        url = f'https://racing.stheadline.com/api/raceOdds/latest?raceNo={race_no}&type=win,place,quin,place-quin&rev=2' 

        response = requests.get(url)

        try:
            wins = response.json()['data']['win']['raceOddsList']

            qins = response.json()['data']['quin']['raceOddsList']

            pd.DataFrame(wins).to_excel(writer, sheet_name=f"win_{race_no}", index=False)
            pd.DataFrame(qins).to_excel(writer, sheet_name=f"qin_{race_no}", index=False)
        except:
            pass

