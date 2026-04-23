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


def add_recent_form(df, n=3):
    """Add '最近N次名次' column to df showing last n finishing positions (newest first).
    Positions are taken from previous races only (excludes current race).
    Horses identified by '馬名'; sorted by '總場次'.
    """
    col = f'最近{n}次名次'
    df = df.sort_values(['馬名', '總場次']).copy()
    records = []
    for _, group in df.groupby('馬名', sort=False):
        group = group.copy()
        positions = group['名次'].tolist()
        form_list = []
        for i in range(len(positions)):
            prev = positions[max(0, i - n):i]
            form_list.append('-'.join(
                str(int(v)) if pd.notna(v) else '?' for v in reversed(prev)
            ))
        group[col] = form_list
        records.append(group)
    return pd.concat(records).sort_values('總場次').reset_index(drop=True)

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------

df_2024 = pd.read_excel('raceresult_2024.xlsx')
df_2025 = pd.read_excel('raceresult_2025.xlsx')
# 欄位: 日期 總場次 馬場 場次 班次 路程 場地狀況 賽事名稱 泥草 賽道 名次 馬號 馬名
#       布號 騎師 練馬師 實際負磅 排位體重 檔位 頭馬距離 沿途走位 完成時間 獨贏賠率 ...

df_2025 = pd.concat([df_2024, df_2025], ignore_index=True)
df_2025 = add_recent_form(df_2025, n=3)

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
            (f'{prefix}近績', '最近3次名次'),
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

# ---------------------------------------------------------------------------
# 各馬匹最近三次出賽名次
# ---------------------------------------------------------------------------

df_14_form = (
    df_2025[df_2025['馬號'].isin(TOP4)]
    [['總場次', '馬場', '班次', '路程', '馬號', '馬名', '騎師', '名次', '最近3次名次']]
    .sort_values(['總場次', '馬號'])
    .copy()
)

# 最近 20 場賽事中 1-4 號馬的出賽紀錄與近況
recent_races = sorted(df_14_form['總場次'].unique())[-20:]
disp_form = df_14_form[df_14_form['總場次'].isin(recent_races)]
print()
print("1-4 號馬最近三次出賽名次 (最近 20 場賽事):")
print(disp_form.to_string(index=False))

# 各馬匹最近3次名次分布統計 (以出賽場次排序)
print()
print("1-4 號馬 近況摘要 (按馬號, 顯示最近 5 次出賽):")
for horse_num in TOP4:
    recent5 = (
        df_2025[df_2025['馬號'] == horse_num]
        .sort_values('總場次', ascending=False)
        .head(5)
        [['總場次', '馬場', '班次', '路程', '馬名', '騎師', '名次', '最近3次名次']]
    )
    print(f"\n  {horse_num} 號馬:")
    print(recent5.to_string(index=False))

# ---------------------------------------------------------------------------
# 沙田草地四班 1200/1400 三甲馬匹近績分析
# ---------------------------------------------------------------------------

def form_has_within_top5(form_str):
    """
    Return True  if at least one position in recent form string is <= 5.
    Return False if all positions are > 5.
    Return None  if no valid form data.
    """
    if not isinstance(form_str, str) or form_str == '':
        return None
    positions = [p for p in form_str.split('-') if p.isdigit()]
    if not positions:
        return None
    return any(int(p) <= 5 for p in positions)


# 篩選條件：沙田 x 草 x 四班 x 1200 或 1400
st_c4 = df_2025[
    (df_2025['馬場'] == 'ST') &
    (df_2025['泥草'] == '草地') &
    (df_2025['班次'] == '第四班 ') &
    (df_2025['路程'].isin([1200, 1400]))
].copy()

# 只取跑入三甲的馬匹
st_c4_top3 = st_c4[st_c4['名次'] <= 3].copy()

st_c4_top3['近績有五名以內'] = st_c4_top3['最近3次名次'].apply(form_has_within_top5)

# 排除無近績資料的馬匹
st_c4_with_form = st_c4_top3[st_c4_top3['近績有五名以內'].notna()].copy()

有五名以內 = (st_c4_with_form['近績有五名以內'] == True).sum()
全五名以外 = (st_c4_with_form['近績有五名以內'] == False).sum()
total = 有五名以內 + 全五名以外

print()
print("=" * 60)
print("沙田草地四班 1200/1400 三甲馬匹近績分析")
print("=" * 60)
print(f"三甲馬匹總數 (有近績資料): {total}")
print(f"  近績包括5名以內 (≥1次名次≤5): {有五名以內} ({有五名以內/total*100:.1f}%)")
print(f"  近績全為5名以外 (全部名次>5):  {全五名以外} ({全五名以外/total*100:.1f}%)")

# 按路程分拆
print()
print("按路程分拆:")
for dist in [1200, 1400]:
    sub = st_c4_with_form[st_c4_with_form['路程'] == dist]
    n_in  = (sub['近績有五名以內'] == True).sum()
    n_out = (sub['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    if n_tot > 0:
        print(f"  {dist}m: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")

# 明細列表
print()
print("三甲馬匹明細 (按總場次):")
print(st_c4_with_form[
    ['總場次', '路程', '名次', '馬號', '馬名', '騎師', '最近3次名次', '近績有五名以內']
].sort_values('總場次').to_string(index=False))

# ---------------------------------------------------------------------------
# 總體三甲馬匹近績分析 (所有馬場 / 班次 / 路程)
# ---------------------------------------------------------------------------

all_top3 = df_2025[df_2025['名次'] <= 3].copy()
all_top3['近績有五名以內'] = all_top3['最近3次名次'].apply(form_has_within_top5)
all_with_form = all_top3[all_top3['近績有五名以內'].notna()].copy()

a_in  = (all_with_form['近績有五名以內'] == True).sum()
a_out = (all_with_form['近績有五名以內'] == False).sum()
a_tot = a_in + a_out

print()
print("=" * 60)
print("總體三甲馬匹近績分析 (所有馬場 / 班次 / 路程)")
print("=" * 60)
print(f"三甲馬匹總數 (有近績資料): {a_tot}")
print(f"  近績包括5名以內 (≥1次名次≤5): {a_in} ({a_in/a_tot*100:.1f}%)")
print(f"  近績全為5名以外 (全部名次>5):  {a_out} ({a_out/a_tot*100:.1f}%)")

# 按馬場 x 泥草 分拆
print()
print("按馬場 x 泥草分拆:")
for (venue, turf), grp in all_with_form.groupby(['馬場', '泥草']):
    n_in  = (grp['近績有五名以內'] == True).sum()
    n_out = (grp['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    print(f"  {venue} {turf}: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")

# 按班次分拆
print()
print("按班次分拆:")
for cls, grp in all_with_form.groupby('班次'):
    n_in  = (grp['近績有五名以內'] == True).sum()
    n_out = (grp['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    print(f"  {cls}: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")

# ---------------------------------------------------------------------------
# 1-4 號馬跑入三甲近績分析
# ---------------------------------------------------------------------------

top4_top3 = df_2025[df_2025['馬號'].isin(TOP4) & (df_2025['名次'] <= 3)].copy()
top4_top3['近績有五名以內'] = top4_top3['最近3次名次'].apply(form_has_within_top5)
top4_with_form = top4_top3[top4_top3['近績有五名以內'].notna()].copy()

t_in  = (top4_with_form['近績有五名以內'] == True).sum()
t_out = (top4_with_form['近績有五名以內'] == False).sum()
t_tot = t_in + t_out

print()
print("=" * 60)
print("1-4 號馬跑入三甲近績分析")
print("=" * 60)
print(f"三甲馬匹總數 (有近績資料): {t_tot}")
print(f"  近績包括5名以內 (≥1次名次≤5): {t_in} ({t_in/t_tot*100:.1f}%)")
print(f"  近績全為5名以外 (全部名次>5):  {t_out} ({t_out/t_tot*100:.1f}%)")

# 按馬號分拆
print()
print("按馬號分拆:")
for horse_num in TOP4:
    grp = top4_with_form[top4_with_form['馬號'] == horse_num]
    n_in  = (grp['近績有五名以內'] == True).sum()
    n_out = (grp['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    if n_tot > 0:
        print(f"  {horse_num}號馬: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")

# 按馬場 x 泥草分拆
print()
print("按馬場 x 泥草分拆:")
for (venue, turf), grp in top4_with_form.groupby(['馬場', '泥草']):
    n_in  = (grp['近績有五名以內'] == True).sum()
    n_out = (grp['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    print(f"  {venue} {turf}: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")

# 按班次分拆
print()
print("按班次分拆:")
for cls, grp in top4_with_form.groupby('班次'):
    n_in  = (grp['近績有五名以內'] == True).sum()
    n_out = (grp['近績有五名以內'] == False).sum()
    n_tot = n_in + n_out
    print(f"  {cls}: 共{n_tot}匹 | 含5名以內 {n_in} ({n_in/n_tot*100:.1f}%) | 全5名以外 {n_out} ({n_out/n_tot*100:.1f}%)")


