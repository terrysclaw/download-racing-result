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

            for win in wins:
                win['qin'] = 0

                for qin in qins:
                    if qin['horseNo1'] == win['horseNo1'] or qin['horseNo2'] == win['horseNo1']:
                        win['qin'] += qin['investment']

            ## rank the horses by the willPay
            wins = sorted(wins, key=lambda x: x['willPay'], reverse=False)

            ## append the rank to the win data
            for i, win in enumerate(wins):
                win['win_rank'] = i + 1

            wins = sorted(wins, key=lambda x: x['qin'], reverse=True)

            for i, win in enumerate(wins):
                win['qin_rank'] = i + 1

                win['win-qin'] = win['win_rank'] - win['qin_rank']

            wins = sorted(wins, key=lambda x: x['horseNo1'], reverse=False)


            pd.DataFrame(wins).to_excel(writer, sheet_name=f"win_{race_no}", index=False)
            # pd.DataFrame(qins).to_excel(writer, sheet_name=f"qin_{race_no}", index=False)
        except:
            pass

