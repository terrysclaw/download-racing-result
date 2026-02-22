import pandas as pd

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

TOP4 = [1, 2, 3, 4]
PLACE_COLS = ['冠軍馬號', '亞軍馬號', '季軍馬號']


def pct_str(series_num, series_den):
    """Return a percentage string Series (e.g. '83.1%')."""
    return (series_num / series_den * 100).round(1).astype(str) + '%'


def top3_count(df_grouped, horse_num):
    """Count appearances of horse_num across all top-3 columns in df_grouped."""
    return sum((df_grouped[col] == horse_num).sum() for col in PLACE_COLS)


def add_horse_counts(stats_df, df_grouped, filter_cols=None):
    """
    Add columns '1號馬'..'4號馬' to stats_df by counting top-3 appearances
    in df_grouped, filtering rows by matching values in filter_cols (list of
    column names that exist in both stats_df and df_grouped).
    Also adds '比率數值' and '最強號碼' convenience columns.
    """
    if filter_cols is None:
        filter_cols = [c for c in stats_df.columns if c in df_grouped.columns
                       and c not in ('總場次', '含前四號馬場次', '比率')]
    for horse_num in TOP4:
        counts = []
        for _, row in stats_df.iterrows():
            mask = pd.Series([True] * len(df_grouped), index=df_grouped.index)
            for col in filter_cols:
                mask &= df_grouped[col] == row[col]
            sub = df_grouped[mask]
            counts.append(sum(int((sub[c] == horse_num).sum()) for c in PLACE_COLS))
        stats_df[f'{horse_num}號馬'] = counts
    stats_df['比率數值'] = stats_df['含前四號馬場次'] / stats_df['總場次'] * 100
    stats_df['最強號碼'] = stats_df.apply(
        lambda r: (lambda d: f"{max(d, key=d.get)}號 ({d[max(d, key=d.get)]}次)")(
            {h: r[f'{h}號馬'] for h in TOP4}
        ), axis=1
    )
    return stats_df


def top4_presence_stats(df_grouped, group_cols, min_races=1):
    """
    Group df_grouped by group_cols, compute 總場次 / 含前四號馬場次 / 比率,
    then attach per-horse counts via add_horse_counts.
    """
    stats = df_grouped.groupby(group_cols).agg(
        總場次=('總場次', 'count'),
        含前四號馬場次=('有前四號馬', 'sum'),
    ).reset_index()
    stats['比率'] = pct_str(stats['含前四號馬場次'], stats['總場次'])
    add_horse_counts(stats, df_grouped, filter_cols=group_cols)
    if min_races > 1:
        stats = stats[stats['總場次'] >= min_races].reset_index(drop=True)
    return stats


def horse_stats_by_group(df_raw, group_cols, min_starts=3):
    """
    For each horse in TOP4, aggregate 出賽次數 / 冠軍 / 前三名 stats grouped by
    group_cols. Returns a combined DataFrame with a '號碼' column.
    """
    rows = []
    for horse_num in TOP4:
        df_h = df_raw[df_raw['馬號'] == horse_num]
        stats = df_h.groupby(group_cols).agg(
            出賽次數=('名次', 'count'),
            冠軍=('名次', lambda x: (x == 1).sum()),
            前三名=('名次', lambda x: (x <= 3).sum()),
        ).reset_index()
        stats['號碼'] = horse_num
        stats['勝率數值'] = stats['冠軍'] / stats['出賽次數'] * 100
        stats['前三名率數值'] = stats['前三名'] / stats['出賽次數'] * 100
        stats['勝率'] = pct_str(stats['冠軍'], stats['出賽次數'])
        stats['前三名率'] = pct_str(stats['前三名'], stats['出賽次數'])
        rows.append(stats[stats['出賽次數'] >= min_starts])
    return pd.concat(rows, ignore_index=True)


def person_stats(df, group_col, min_starts=5):
    """Aggregate 出賽/冠軍/亞軍/季軍/前三名 stats for a jockey or trainer column."""
    stats = df.groupby(group_col).agg(
        出賽次數=('名次', 'count'),
        冠軍次數=('名次', lambda x: (x == 1).sum()),
        亞軍次數=('名次', lambda x: (x == 2).sum()),
        季軍次數=('名次', lambda x: (x == 3).sum()),
        前三名次數=('名次', lambda x: (x <= 3).sum()),
    ).reset_index()
    stats['勝率數值'] = stats['冠軍次數'] / stats['出賽次數'] * 100
    stats['勝率'] = pct_str(stats['冠軍次數'], stats['出賽次數'])
    stats['前三名率'] = pct_str(stats['前三名次數'], stats['出賽次數'])
    return stats[stats['出賽次數'] >= min_starts].sort_values('勝率數值', ascending=False)


PERSON_COLS = ['出賽次數', '冠軍次數', '亞軍次數', '季軍次數', '前三名次數', '勝率', '前三名率']
HORSE_COLS4 = ['總場次', '含前四號馬場次', '比率', '1號馬', '2號馬', '3號馬', '4號馬']

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------

df_2025 = pd.read_excel('raceresult_2025.xlsx')
# 欄位: 日期 總場次 馬場 場次 班次 路程 場地狀況 賽事名稱 泥草 賽道 名次 馬號 馬名
#       布號 騎師 練馬師 實際負磅 排位體重 檔位 頭馬距離 沿途走位 完成時間 獨贏賠率 ...

# ---------------------------------------------------------------------------
# Build race-level summary (one row per race)
# ---------------------------------------------------------------------------

def _first(series):
    return series.values[0] if not series.empty else None

df_grouped = df_2025.groupby(['總場次', '場次', '馬場', '泥草', '班次', '路程']).apply(
    lambda x: pd.Series({
        place: _first(x.loc[x['名次'] == rank, col])
        for rank, prefix in zip([1, 2, 3], ['冠軍', '亞軍', '季軍'])
        for place, col in [
            (f'{prefix}馬號', '馬號'), (f'{prefix}馬名', '馬名'),
            (f'{prefix}賠率', '獨贏賠率'), (f'{prefix}騎師', '騎師'),
        ]
    })
).reset_index()

try:
    df_grouped.to_excel('grouped_results_2025.xlsx', index=False)
except PermissionError:
    print("Warning: grouped_results_2025.xlsx is open, skipping Excel export.")

# ---------------------------------------------------------------------------
# 目標騎師分析
# ---------------------------------------------------------------------------

TARGET_JOCKEYS = [
    '梁家俊', '蔡明紹', '何澤堯', '鍾易禮', '潘明輝',
    '周俊樂', '楊明綸', '黃智弘', '黃寶妮',
]

df_grouped['目標騎師'] = df_grouped.apply(
    lambda row: ', '.join(
        j for j in [row.get('冠軍騎師'), row.get('亞軍騎師'), row.get('季軍騎師')]
        if j in TARGET_JOCKEYS
    ), axis=1
)
df_with_target = df_grouped[df_grouped['目標騎師'] != ''].reset_index(drop=True)

print(f"共 {len(df_grouped)} 場賽事")
print(f"前三名包含目標騎師的場次: {len(df_with_target)}")
print()
print("目標騎師前三名出現次數:")
for j in TARGET_JOCKEYS:
    count = sum((df_grouped[col] == j).sum() for col in ['冠軍騎師', '亞軍騎師', '季軍騎師'])
    print(f"  {j}: {count} 次")
print()
print(df_with_target[['總場次', '班次', '路程', '冠軍騎師', '亞軍騎師', '季軍騎師', '目標騎師']].to_string(index=False))

# ---------------------------------------------------------------------------
# 1-4 號馬前三名分析
# ---------------------------------------------------------------------------

df_grouped['前四號馬'] = df_grouped.apply(
    lambda row: ', '.join(
        str(int(h)) for h in [row.get(c) for c in PLACE_COLS]
        if h is not None and pd.notna(h) and int(h) in TOP4
    ), axis=1
)
df_grouped['有前四號馬'] = df_grouped['前四號馬'] != ''

# ---------------------------------------------------------------------------
# 各維度統計
# ---------------------------------------------------------------------------

class_stats = top4_presence_stats(df_grouped, ['班次'])
print()
print("各班次 1-4 號馬前三名統計:")
print(class_stats[['班次'] + HORSE_COLS4].to_string(index=False))

venue_turf_stats = top4_presence_stats(df_grouped, ['馬場', '泥草'])
print()
print("各馬場 x 泥草 1-4 號馬前三名統計:")
print(venue_turf_stats[['馬場', '泥草'] + HORSE_COLS4].to_string(index=False))

full_combo = top4_presence_stats(df_grouped, ['馬場', '泥草', '班次', '路程'])
print()
print("馬場 x 泥草 x 班次 x 路程 1-4 號馬前三名統計 (總場次>=3):")
disp = full_combo[full_combo['總場次'] >= 3].sort_values(['馬場', '泥草', '班次', '路程'])
print(disp[['馬場', '泥草', '班次', '路程'] + HORSE_COLS4].to_string(index=False))

# ---------------------------------------------------------------------------
# 策騎 1-4 號馬的騎師統計
# ---------------------------------------------------------------------------

df_14 = df_2025[df_2025['馬號'].isin(TOP4)].copy()
jockey_14 = df_14.groupby('騎師').agg(
    出賽次數=('名次', 'count'),
    前三名次數=('名次', lambda x: (x <= 3).sum()),
    冠軍次數=('名次', lambda x: (x == 1).sum()),
).reset_index()
jockey_14['前三名率數值'] = jockey_14['前三名次數'] / jockey_14['出賽次數'] * 100
jockey_14['勝率'] = pct_str(jockey_14['冠軍次數'], jockey_14['出賽次數'])
jockey_14['前三名率'] = pct_str(jockey_14['前三名次數'], jockey_14['出賽次數'])

jockey_14_top = jockey_14[jockey_14['出賽次數'] >= 10].sort_values('前三名率數值', ascending=False)
print()
print("策騎 1-4 號馬騎師統計 (出賽次數>=10, 按前三名率排序):")
print(jockey_14_top[['騎師', '出賽次數', '冠軍次數', '勝率', '前三名次數', '前三名率']].to_string(index=False))


