import pandas as pd


df_2025 = pd.read_excel('raceresult_2025.xlsx')
# 日期	總場次	馬場	場次	班次	路程	場地狀況	賽事名稱	泥草	賽道	賽事時間1	賽事時間2	賽事時間3	賽事時間4	賽事時間5	賽事時間6	名次	馬號	馬名	布號	騎師	練馬師	實際負磅	排位體重	檔位	頭馬距離	沿途走位	完成時間	獨贏賠率	第 1 段	第 2 段	第 3 段	第 4 段	第 5 段	第 6 段

## output group by 總場次, 班次, 路程, 冠軍馬號, 冠軍馬名, 冠軍賠率, 亞軍馬號, 亞軍馬名, 亞軍賠率, 季軍馬號, 季軍馬名, 季軍賠率, 騎師

df_grouped = df_2025.groupby(['總場次', '場次', '馬場', '泥草', '班次', '路程']).apply(lambda x: pd.Series({
    '冠軍馬號': x.loc[x['名次'] == 1, '馬號'].values[0] if not x.loc[x['名次'] == 1, '馬號'].empty else None,
    '冠軍馬名': x.loc[x['名次'] == 1, '馬名'].values[0] if not x.loc[x['名次'] == 1, '馬名'].empty else None,
    '冠軍賠率': x.loc[x['名次'] == 1, '獨贏賠率'].values[0] if not x.loc[x['名次'] == 1, '獨贏賠率'].empty else None,
    '冠軍騎師': x.loc[x['名次'] == 1, '騎師'].values[0] if not x.loc[x['名次'] == 1, '騎師'].empty else None,
    '亞軍馬號': x.loc[x['名次'] == 2, '馬號'].values[0] if not x.loc[x['名次'] == 2, '馬號'].empty else None,
    '亞軍馬名': x.loc[x['名次'] == 2, '馬名'].values[0] if not x.loc[x['名次'] == 2, '馬名'].empty else None,
    '亞軍賠率': x.loc[x['名次'] == 2, '獨贏賠率'].values[0] if not x.loc[x['名次'] == 2, '獨贏賠率'].empty else None,
    '亞軍騎師': x.loc[x['名次'] == 2, '騎師'].values[0] if not x.loc[x['名次'] == 2, '騎師'].empty else None,
    '季軍馬號': x.loc[x['名次'] == 3, '馬號'].values[0] if not x.loc[x['名次'] == 3, '馬號'].empty else None,
    '季軍馬名': x.loc[x['名次'] == 3, '馬名'].values[0] if not x.loc[x['名次'] == 3, '馬名'].empty else None,
    '季軍賠率': x.loc[x['名次'] == 3, '獨贏賠率'].values[0] if not x.loc[x['名次'] == 3, '獨贏賠率'].empty else None,
    '季軍騎師': x.loc[x['名次'] == 3, '騎師'].values[0] if not x.loc[x['名次'] == 3, '騎師'].empty else None,
})).reset_index()



## output to excel
try:
    df_grouped.to_excel('grouped_results_2025.xlsx', index=False)
except PermissionError:
    print("Warning: grouped_results_2025.xlsx is open, skipping Excel export.")


## 檢查每場賽事前三名的騎師是否包括以下騎師
target_jockeys = [
    '梁家俊', '蔡明紹', '何澤堯', '鍾易禮', '潘明輝',
    '周俊樂', '楊明綸', '黃智弘', '黃寶妮'
]

def check_jockeys(row):
    top3 = [row.get('冠軍騎師'), row.get('亞軍騎師'), row.get('季軍騎師')]
    matched = [j for j in top3 if j in target_jockeys]
    return ', '.join(matched) if matched else ''

df_grouped['目標騎師'] = df_grouped.apply(check_jockeys, axis=1)
df_with_target = df_grouped[df_grouped['目標騎師'] != ''].reset_index(drop=True)

print(f"共 {len(df_grouped)} 場賽事")
print(f"前三名包含目標騎師的場次: {len(df_with_target)}")
print()

## 統計每位目標騎師在前三名出現的次數
jockey_counts = {}
for j in target_jockeys:
    count = (
        (df_grouped['冠軍騎師'] == j).sum() +
        (df_grouped['亞軍騎師'] == j).sum() +
        (df_grouped['季軍騎師'] == j).sum()
    )
    jockey_counts[j] = count

print("目標騎師前三名出現次數:")
for j, c in jockey_counts.items():
    print(f"  {j}: {c} 次")

print()
print(df_with_target[['總場次', '班次', '路程', '冠軍騎師', '亞軍騎師', '季軍騎師', '目標騎師']].to_string(index=False))


## 檢查每場賽事前三名的馬匹是否 1-4 號馬
def check_top4_horses(row):
    top3 = [row.get('冠軍馬號'), row.get('亞軍馬號'), row.get('季軍馬號')]
    matched = [str(int(h)) for h in top3 if h is not None and pd.notna(h) and int(h) in [1, 2, 3, 4]]
    return ', '.join(matched) if matched else ''

df_grouped['前四號馬'] = df_grouped.apply(check_top4_horses, axis=1)
df_with_top4 = df_grouped[df_grouped['前四號馬'] != ''].reset_index(drop=True)

print()
print(f"前三名包含 1-4 號馬的場次: {len(df_with_top4)} / {len(df_grouped)}")
print()

## 統計 1-4 號馬在前三名出現的次數
for horse_num in [1, 2, 3, 4]:
    count = (
        (df_grouped['冠軍馬號'] == horse_num).sum() +
        (df_grouped['亞軍馬號'] == horse_num).sum() +
        (df_grouped['季軍馬號'] == horse_num).sum()
    )
    print(f"  {horse_num} 號馬前三名出現次數: {count} 次")

print()
print(df_with_top4[['總場次', '班次', '路程', '冠軍馬號', '亞軍馬號', '季軍馬號', '前四號馬']].to_string(index=False))


## 計算各路程包含 1-4 號馬在前三名的比率
def has_top4(row):
    top3 = [row.get('冠軍馬號'), row.get('亞軍馬號'), row.get('季軍馬號')]
    return any(h is not None and pd.notna(h) and int(h) in [1, 2, 3, 4] for h in top3)

df_grouped['有前四號馬'] = df_grouped.apply(has_top4, axis=1)

dist_stats = df_grouped.groupby('路程').agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
dist_stats['比率'] = (dist_stats['含前四號馬場次'] / dist_stats['總場次'] * 100).round(1).astype(str) + '%'

## 各號馬各路程出現次數
for horse_num in [1, 2, 3, 4]:
    counts = []
    for dist in dist_stats['路程']:
        sub = df_grouped[df_grouped['路程'] == dist]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    dist_stats[f'{horse_num}號馬'] = counts

print()
print("各路程 1-4 號馬前三名統計:")
print(dist_stats[['路程', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 計算各班次包含 1-4 號馬在前三名的比率
class_stats = df_grouped.groupby('班次').agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
class_stats['比率'] = (class_stats['含前四號馬場次'] / class_stats['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for cls in class_stats['班次']:
        sub = df_grouped[df_grouped['班次'] == cls]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    class_stats[f'{horse_num}號馬'] = counts

print()
print("各班次 1-4 號馬前三名統計:")
print(class_stats[['班次', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 結合班次及路程，計算包含 1-4 號馬比率
combo_stats = df_grouped.groupby(['班次', '路程']).agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
combo_stats['比率'] = (combo_stats['含前四號馬場次'] / combo_stats['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for _, row in combo_stats.iterrows():
        sub = df_grouped[(df_grouped['班次'] == row['班次']) & (df_grouped['路程'] == row['路程'])]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    combo_stats[f'{horse_num}號馬'] = counts

print()
print("班次 x 路程 1-4 號馬前三名統計:")
print(combo_stats[['班次', '路程', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 計算各馬場包含 1-4 號馬在前三名的比率
venue_stats = df_grouped.groupby('馬場').agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
venue_stats['比率'] = (venue_stats['含前四號馬場次'] / venue_stats['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for v in venue_stats['馬場']:
        sub = df_grouped[df_grouped['馬場'] == v]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    venue_stats[f'{horse_num}號馬'] = counts

print()
print("各馬場 1-4 號馬前三名統計:")
print(venue_stats[['馬場', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 計算各馬場 x 泥草包含 1-4 號馬在前三名的比率
venue_turf_stats = df_grouped.groupby(['馬場', '泥草']).agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
venue_turf_stats['比率'] = (venue_turf_stats['含前四號馬場次'] / venue_turf_stats['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for _, row in venue_turf_stats.iterrows():
        sub = df_grouped[(df_grouped['馬場'] == row['馬場']) & (df_grouped['泥草'] == row['泥草'])]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    venue_turf_stats[f'{horse_num}號馬'] = counts

print()
print("各馬場 x 泥草 1-4 號馬前三名統計:")
print(venue_turf_stats[['馬場', '泥草', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 結合馬場, 泥草, 班次及路程，計算包含 1-4 號馬比率
full_combo = df_grouped.groupby(['馬場', '泥草', '班次', '路程']).agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
full_combo['比率'] = (full_combo['含前四號馬場次'] / full_combo['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for _, row in full_combo.iterrows():
        sub = df_grouped[
            (df_grouped['馬場'] == row['馬場']) &
            (df_grouped['泥草'] == row['泥草']) &
            (df_grouped['班次'] == row['班次']) &
            (df_grouped['路程'] == row['路程'])
        ]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    full_combo[f'{horse_num}號馬'] = counts

print()
print("馬場 x 泥草 x 班次 x 路程 1-4 號馬前三名統計 (總場次>=3):")
display = full_combo[full_combo['總場次'] >= 3].sort_values(['馬場', '泥草', '班次', '路程'])
print(display[['馬場', '泥草', '班次', '路程', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 草地賽事最值得投注 1-4 號馬的組合 (總場次>=5, 按比率排序)
full_combo['比率數值'] = full_combo['含前四號馬場次'] / full_combo['總場次'] * 100
turf_best = (
    full_combo[(full_combo['泥草'] == '草地') & (full_combo['總場次'] >= 5)]
    .sort_values('比率數值', ascending=False)
)

print()
print("草地賽事最值得投注 1-4 號馬的組合 (總場次>=5, 按比率排序):")
print(turf_best[['馬場', '班次', '路程', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].to_string(index=False))


## 統計各場次包含 1-4 號馬比率
race_num_stats = df_grouped.groupby('場次').agg(
    總場次=('總場次', 'count'),
    含前四號馬場次=('有前四號馬', 'sum')
).reset_index()
race_num_stats['比率'] = (race_num_stats['含前四號馬場次'] / race_num_stats['總場次'] * 100).round(1).astype(str) + '%'

for horse_num in [1, 2, 3, 4]:
    counts = []
    for rn in race_num_stats['場次']:
        sub = df_grouped[df_grouped['場次'] == rn]
        counts.append(
            int((sub['冠軍馬號'] == horse_num).sum()) +
            int((sub['亞軍馬號'] == horse_num).sum()) +
            int((sub['季軍馬號'] == horse_num).sum())
        )
    race_num_stats[f'{horse_num}號馬'] = counts

print()
print("各場次 1-4 號馬前三名統計:")
print(race_num_stats[['場次', '總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']].sort_values('場次').to_string(index=False))


## 沙田草地第四班1200米 - 騎師及練馬師成績
df_st_c4_1200 = df_2025[
    (df_2025['馬場'] == 'ST') &
    (df_2025['泥草'].str.contains('草地', na=False)) &
    (df_2025['班次'].str.contains('第四班', na=False)) &
    (~df_2025['班次'].str.contains('條件', na=False)) &
    (df_2025['路程'] == 1200)
].copy()

print()
print(f"=== 沙田草地第四班1200米 ===")
print(f"共 {df_st_c4_1200['總場次'].nunique()} 場賽事, {len(df_st_c4_1200)} 條出賽記錄")

## 騎師成績
jockey_st_c4_1200 = df_st_c4_1200.groupby('騎師').agg(
    出賽次數=('名次', 'count'),
    冠軍次數=('名次', lambda x: (x == 1).sum()),
    亞軍次數=('名次', lambda x: (x == 2).sum()),
    季軍次數=('名次', lambda x: (x == 3).sum()),
    前三名次數=('名次', lambda x: (x <= 3).sum()),
).reset_index()

jockey_st_c4_1200['勝率'] = (jockey_st_c4_1200['冠軍次數'] / jockey_st_c4_1200['出賽次數'] * 100).round(1).astype(str) + '%'
jockey_st_c4_1200['前三名率'] = (jockey_st_c4_1200['前三名次數'] / jockey_st_c4_1200['出賽次數'] * 100).round(1).astype(str) + '%'
jockey_st_c4_1200['勝率數值'] = jockey_st_c4_1200['冠軍次數'] / jockey_st_c4_1200['出賽次數'] * 100

jockey_st_c4_1200_top = jockey_st_c4_1200[jockey_st_c4_1200['出賽次數'] >= 5].sort_values('勝率數值', ascending=False)

print()
print("騎師成績 (出賽次數>=5, 按勝率排序):")
print(jockey_st_c4_1200_top[['騎師', '出賽次數', '冠軍次數', '亞軍次數', '季軍次數', '前三名次數', '勝率', '前三名率']].to_string(index=False))

## 練馬師成績
trainer_st_c4_1200 = df_st_c4_1200.groupby('練馬師').agg(
    出賽次數=('名次', 'count'),
    冠軍次數=('名次', lambda x: (x == 1).sum()),
    亞軍次數=('名次', lambda x: (x == 2).sum()),
    季軍次數=('名次', lambda x: (x == 3).sum()),
    前三名次數=('名次', lambda x: (x <= 3).sum()),
).reset_index()

trainer_st_c4_1200['勝率'] = (trainer_st_c4_1200['冠軍次數'] / trainer_st_c4_1200['出賽次數'] * 100).round(1).astype(str) + '%'
trainer_st_c4_1200['前三名率'] = (trainer_st_c4_1200['前三名次數'] / trainer_st_c4_1200['出賽次數'] * 100).round(1).astype(str) + '%'
trainer_st_c4_1200['勝率數值'] = trainer_st_c4_1200['冠軍次數'] / trainer_st_c4_1200['出賽次數'] * 100

trainer_st_c4_1200_top = trainer_st_c4_1200[trainer_st_c4_1200['出賽次數'] >= 5].sort_values('勝率數值', ascending=False)

print()
print("練馬師成績 (出賽次數>=5, 按勝率排序):")
print(trainer_st_c4_1200_top[['練馬師', '出賽次數', '冠軍次數', '亞軍次數', '季軍次數', '前三名次數', '勝率', '前三名率']].to_string(index=False))


## 那個騎師策騎的 1-4 號馬最好
# 從原始資料篩選馬號 1-4 的出賽記錄
df_14 = df_2025[df_2025['馬號'].isin([1, 2, 3, 4])].copy()

jockey_14 = df_14.groupby('騎師').agg(
    出賽次數=('名次', 'count'),
    前三名次數=('名次', lambda x: (x <= 3).sum()),
    冠軍次數=('名次', lambda x: (x == 1).sum()),
).reset_index()

jockey_14['前三名率'] = (jockey_14['前三名次數'] / jockey_14['出賽次數'] * 100).round(1).astype(str) + '%'
jockey_14['勝率'] = (jockey_14['冠軍次數'] / jockey_14['出賽次數'] * 100).round(1).astype(str) + '%'
jockey_14['前三名率數值'] = jockey_14['前三名次數'] / jockey_14['出賽次數'] * 100

# 只顯示出賽次數 >= 10 的騎師，按前三名率排序
jockey_14_top = jockey_14[jockey_14['出賽次數'] >= 10].sort_values('前三名率數值', ascending=False)

print()
print("策騎 1-4 號馬騎師統計 (出賽次數>=10, 按前三名率排序):")
print(jockey_14_top[['騎師', '出賽次數', '冠軍次數', '勝率', '前三名次數', '前三名率']].to_string(index=False))


