import pandas as pd
import re
import os
import glob
import warnings

warnings.filterwarnings('ignore')


# ==================== 工具函数 ====================
def safe_to_numeric(series, default_value=0):
    result = pd.to_numeric(series, errors='coerce')
    return result.fillna(default_value).astype(float)


def extract_furnace_level(level_str):
    if pd.isna(level_str) or str(level_str).strip() == '':
        return 0
    level_str = str(level_str).strip()
    match = re.search(r'lv\.?\s*(\d+)', level_str, re.IGNORECASE)
    if match:
        level = int(match.group(1))
        return level if level <= MAX_LEVEL else 0
    return 0


def has_collecting(troop_str):
    if pd.isna(troop_str) or str(troop_str).strip() == '':
        return False
    return '采集资源' in str(troop_str)


def count_collecting_teams(troop_str):
    if pd.isna(troop_str) or str(troop_str).strip() == '':
        return 0
    return str(troop_str).count('采集资源')


# ==================== 配置 ====================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_PATH = os.path.join(CURRENT_DIR, 'data')
FILE_PATTERN = '*.xlsx'
THRESHOLD = 0.5
MAX_PRESTIGE = 2000
MAX_LEVEL = 30

# ==================== 查找文件 ====================
all_files = glob.glob(os.path.join(FOLDER_PATH, FILE_PATTERN))
all_files = sorted(all_files)
print(f"找到 {len(all_files)} 个Excel文件:")
for f in all_files:
    print(f"  - {os.path.basename(f)}")

if len(all_files) == 0:
    print("❌ 没有找到xlsx文件")
    exit()

# ==================== 读取所有文件 ====================
daily_accounts = []

for file_path in all_files:
    filename = os.path.basename(file_path)
    df = pd.read_excel(file_path, sheet_name=0, dtype=str)

    # 🔧 检查是否有"兵线"列，没有则跳过
    if '兵线' not in df.columns:
        print(f"  ⚠️ {filename}: 无'兵线'列，跳过")
        continue

    df['账号'] = safe_to_numeric(df['账号'])
    df['声望'] = safe_to_numeric(df['声望'])
    df['炉子等级'] = df['炉子'].apply(extract_furnace_level)

    today_dict = {}

    for _, row in df.iterrows():
        account = int(row['账号'])
        if account == 0:
            continue

        if row['炉子等级'] < 1 or row['炉子等级'] > MAX_LEVEL:
            continue
        if row['声望'] > MAX_PRESTIGE:
            continue

        is_mining = has_collecting(row['兵线'])
        teams = count_collecting_teams(row['兵线'])

        today_dict[account] = {
            'is_mining': is_mining,
            'teams': teams,
            'info': {
                '昵称': row['昵称'],
                '联盟': row['联盟'],
                '炉子': row['炉子'],
                '炉子等级': int(row['炉子等级']),
                '声望': int(row['声望']),
                '坐标': row.get('坐标', ''),
                '罩子': row.get('罩子', ''),
                '兵线': row.get('兵线', ''),
            }
        }

    daily_accounts.append(today_dict)
    mining_count = sum(1 for v in today_dict.values() if v['is_mining'])
    print(f"  {filename}: {len(df)}行 → 条件内{len(today_dict)}人, 挖矿{mining_count}人")

# ==================== 汇总 ====================
total_days = len(daily_accounts)

all_accounts = set()
for day_dict in daily_accounts:
    all_accounts.update(day_dict.keys())

miners = []

for account in all_accounts:
    mining_days = 0
    max_teams = 0
    latest_info = None

    for day_dict in daily_accounts:
        if account in day_dict:
            if day_dict[account]['is_mining']:
                mining_days += 1
            max_teams = max(max_teams, day_dict[account]['teams'])
            latest_info = day_dict[account]['info']

    ratio = mining_days / total_days

    if ratio >= THRESHOLD:
        miners.append({
            '账号': account,
            '昵称': latest_info['昵称'] if latest_info else '',
            '联盟': latest_info['联盟'] if latest_info else '',
            '炉子': latest_info['炉子'] if latest_info else '',
            '炉子等级': latest_info['炉子等级'] if latest_info else 0,
            '声望': latest_info['声望'] if latest_info else 0,
            '坐标': latest_info['坐标'] if latest_info else '',
            '罩子': latest_info['罩子'] if latest_info else '',
            '最大采集队数': max_teams,
            '挖矿天数': f"{mining_days}/{total_days}",
            '挖矿比例': round(ratio, 2),
            '_ratio': ratio,
            '_mining_days': mining_days,
        })

df_miners = pd.DataFrame(miners)

if len(df_miners) > 0:
    df_miners = df_miners.sort_values(['_ratio', '_mining_days', '声望'], ascending=[False, False, False])
    df_miners = df_miners.drop(columns=['_ratio', '_mining_days'])
    df_miners = df_miners.reset_index(drop=True)
    df_miners.index = df_miners.index + 1
    df_miners.index.name = '序号'
    # 挖矿比例格式化
    df_miners['挖矿比例'] = df_miners['挖矿比例'].apply(lambda x: f'{x:.0%}')
    print(f"\n✅ 高频挖矿小号: {len(df_miners)} 个")
else:
    print(f"\n❌ 没有找到满足条件的玩家")

# ==================== 保存 ====================
output_file = f'高频挖矿小号_{total_days}天分析.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    if len(df_miners) > 0:
        output_cols = ['账号', '昵称', '联盟', '炉子', '炉子等级', '声望',
                       '坐标', '罩子', '最大采集队数', '挖矿天数', '挖矿比例']
        df_miners[output_cols].to_excel(writer, sheet_name='高频挖矿小号', index=True)

print(f"\n📁 结果已保存: {output_file}")

if len(df_miners) > 0:
    print(f"\n🔥 高频挖矿小号（前20）:")
    for i, (_, row) in enumerate(df_miners.head(20).iterrows()):
        print(
            f"  {i + 1}. [{row['联盟']}] {row['昵称']} | {row['炉子']} | 声望{row['声望']} | {row['最大采集队数']}队 | {row['挖矿天数']}")