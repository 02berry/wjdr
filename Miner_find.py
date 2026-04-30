import pandas as pd
import re
import os
import glob
import warnings
from datetime import datetime
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

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


def extract_date_from_filename(filename):
    """
    从文件名中提取日期信息
    支持格式：
    - 3957_0428 (一天一次)
    - 3957_0428a, 3957_0428b (一天多次，用abcd区分)
    返回 (datetime对象, 日期字符串, 序号) 用于排序
    """
    # 匹配模式：3957_后面4位数字+可选字母
    match = re.search(r'3957_(\d{4})([a-z]?)', filename)
    if match:
        date_str = match.group(1)  # 0428
        suffix = match.group(2)  # a/b/c 或空字符串

        # 尝试解析日期（假设是当前年份）
        current_year = datetime.now().year
        try:
            month = int(date_str[:2])
            day = int(date_str[2:])
            date_obj = datetime(current_year, month, day)
            return date_obj, date_str, suffix
        except ValueError:
            pass

        # 如果上面失败，尝试6位格式 MMDDYY
        match2 = re.search(r'3957_(\d{6})([a-z]?)', filename)
        if match2:
            date_str = match2.group(1)
            suffix = match2.group(2)
            try:
                year = 2000 + int(date_str[4:6])
                month = int(date_str[:2])
                day = int(date_str[2:4])
                date_obj = datetime(year, month, day)
                return date_obj, date_str, suffix
            except ValueError:
                pass

    return None, None, ''


def load_remark_data(folder_path):
    """从data文件夹加载miner_remark.xlsx备注数据"""
    remark_file = os.path.join(folder_path, 'miner_remark.xlsx')

    if not os.path.exists(remark_file):
        print(f"  ℹ️  未找到备注文件: {remark_file}，将跳过备注功能")
        return {}

    try:
        remark_df = pd.read_excel(remark_file, dtype=str)

        # 检查必要的列
        if '账号' not in remark_df.columns or '备注' not in remark_df.columns:
            print(f"  ⚠️  备注文件格式错误，需要包含'账号'和'备注'两列")
            return {}

        # 转换为字典 {账号: 备注}
        remark_dict = {}
        for _, row in remark_df.iterrows():
            account = row['账号']
            note = row['备注']

            # 尝试将账号转换为数字（如果可能的话）
            try:
                account = int(float(account))
            except (ValueError, TypeError):
                account = str(account).strip()

            if pd.notna(note) and str(note).strip() != '':
                remark_dict[account] = str(note).strip()

        print(f"  ✅ 加载了 {len(remark_dict)} 条备注信息")
        return remark_dict

    except Exception as e:
        print(f"  ❌ 读取备注文件失败: {e}")
        return {}


# ==================== 配置区 ====================
# 【修改文件读取数量】将下面的数字改成你想要读取的最近文件数量
READ_FILE_COUNT = 10  # 读取最近的文件数量，例如改为5则读取最新5个文件

# 其他配置
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_PATH = os.path.join(CURRENT_DIR, 'data')
FILE_PATTERN = '3957_*.xlsx'  # 文件命名模式：3957_0428.xlsx 或 3957_0428a.xlsx
THRESHOLD = 0.5
MAX_PRESTIGE = 2000
MAX_LEVEL = 30

# ==================== 加载备注数据 ====================
print("📋 加载备注数据...")
WATCHED_ACCOUNTS = load_remark_data(FOLDER_PATH)

# ==================== 查找并筛选文件 ====================
all_files = glob.glob(os.path.join(FOLDER_PATH, FILE_PATTERN))

# 过滤掉备注文件本身
all_files = [f for f in all_files if 'miner_remark' not in os.path.basename(f).lower()]

print(f"\n📁 找到 {len(all_files)} 个数据文件")

# 提取文件日期并排序
file_dates = []
for file_path in all_files:
    filename = os.path.basename(file_path)
    date_obj, date_str, suffix = extract_date_from_filename(filename)
    if date_obj:
        # 排序优先级：1.日期 2.后缀字母（同一天内a最早，b其次）
        file_dates.append((file_path, date_obj, date_str, suffix))
    else:
        print(f"  ⚠️ 无法提取日期: {filename}")

# 按日期排序（最新的在前），同一天按后缀字母排序
file_dates.sort(key=lambda x: (x[1], x[3] if x[3] else ''), reverse=False)
file_dates.sort(key=lambda x: x[1], reverse=True)

# 统计实际天数（不同日期的数量）
unique_dates = set(date_str for _, _, date_str, _ in file_dates)
print(f"  📅 涵盖 {len(unique_dates)} 个不同日期")

# 只保留最近的READ_FILE_COUNT天（不是文件数）
sorted_dates = sorted(unique_dates, reverse=True)[:READ_FILE_COUNT]
selected_files = [(fp, do, ds, sf) for fp, do, ds, sf in file_dates if ds in sorted_dates]

print(f"\n✅ 将处理最近 {len(sorted_dates)} 天的数据（共{len(selected_files)}个文件）:")
current_date = None
for file_path, date_obj, date_str, suffix in selected_files:
    display_date = f"{date_str[:2]}-{date_str[2:]}"
    if suffix:
        display_name = f"{display_date} (第{suffix.upper()}次)"
    else:
        display_name = f"{display_date}"

    # 显示日期分组
    if date_str != current_date:
        print(f"  📅 {display_date}:")
        current_date = date_str

    print(f"    📄 {os.path.basename(file_path)}")

if len(selected_files) == 0:
    print("❌ 没有找到符合条件的文件")
    exit()

# ==================== 读取选定文件 ====================
daily_accounts = []
# ⭐ 记录每个账号的最新完整信息（不受条件限制）
account_latest_full_info = {}  # {账号: (日期排序键, info字典)}

for file_path, date_obj, date_str, suffix in selected_files:
    filename = os.path.basename(file_path)

    try:
        df = pd.read_excel(file_path, sheet_name=0, dtype=str)
    except Exception as e:
        print(f"  ❌ 读取失败 {filename}: {e}")
        continue

    if '兵线' not in df.columns:
        print(f"  ⚠️ {filename}: 无'兵线'列，跳过")
        continue

    df['账号'] = safe_to_numeric(df['账号'])
    df['声望'] = safe_to_numeric(df['声望'])
    df['炉子等级'] = df['炉子'].apply(extract_furnace_level)

    today_dict = {}
    suffix_order = ord(suffix) - ord('a') + 1 if suffix else 0
    date_key = (date_obj, suffix_order)

    for _, row in df.iterrows():
        account = int(row['账号'])
        if account == 0:
            continue

        # ⭐ 不管符不符合条件，都记录最新完整信息
        current_info = {
            '昵称': row['昵称'],
            '联盟': row['联盟'],
            '炉子': row['炉子'],
            '炉子等级': int(row['炉子等级']),
            '声望': int(row['声望']),
            '坐标': row.get('坐标', ''),
            '罩子': row.get('罩子', ''),
            '兵线': row.get('兵线', ''),
        }

        if account not in account_latest_full_info:
            account_latest_full_info[account] = (date_key, current_info)
        else:
            if date_key > account_latest_full_info[account][0]:
                account_latest_full_info[account] = (date_key, current_info)

        # 条件过滤：判断是否算作挖矿账号
        if row['炉子等级'] < 1 or row['炉子等级'] > MAX_LEVEL:
            continue
        if row['声望'] > MAX_PRESTIGE:
            continue

        is_mining = has_collecting(row['兵线'])
        teams = count_collecting_teams(row['兵线'])

        today_dict[account] = {
            'is_mining': is_mining,
            'teams': teams,
            'info': current_info,  # 这里也用同样的info
        }

    daily_accounts.append(today_dict)
    mining_count = sum(1 for v in today_dict.values() if v['is_mining'])

    display_date = f"{date_str[:2]}-{date_str[2:]}"
    print(
        f"  {display_date}{'(' + suffix.upper() + ')' if suffix else ''}: {len(df)}行 → {len(today_dict)}人, 挖矿{mining_count}人")

# ==================== 汇总 ====================
total_days = len(daily_accounts)

if total_days == 0:
    print("❌ 没有有效数据可以分析")
    exit()

all_accounts = set()
for day_dict in daily_accounts:
    all_accounts.update(day_dict.keys())

miners = []

for account in all_accounts:
    mining_days = 0
    max_teams = 0
    latest_mining_info = None

    for day_dict in daily_accounts:
        if account in day_dict:
            if day_dict[account]['is_mining']:
                mining_days += 1
            max_teams = max(max_teams, day_dict[account]['teams'])
            latest_mining_info = day_dict[account]['info']

    ratio = mining_days / total_days

    if ratio >= THRESHOLD:
        note = WATCHED_ACCOUNTS.get(account, '')

        # ⭐ 使用最新完整信息（不受挖矿条件限制）
        if account in account_latest_full_info:
            latest_full_info = account_latest_full_info[account][1]
            final_nickname = latest_full_info['昵称']
            final_alliance = latest_full_info['联盟']
            final_furnace = latest_full_info['炉子']
            final_furnace_level = latest_full_info['炉子等级']
            final_prestige = latest_full_info['声望']
            final_coord = latest_full_info['坐标']
            final_shield = latest_full_info['罩子']
        else:
            # 兜底：使用最后的挖矿信息
            final_nickname = latest_mining_info['昵称'] if latest_mining_info else ''
            final_alliance = latest_mining_info['联盟'] if latest_mining_info else ''
            final_furnace = latest_mining_info['炉子'] if latest_mining_info else ''
            final_furnace_level = latest_mining_info['炉子等级'] if latest_mining_info else 0
            final_prestige = latest_mining_info['声望'] if latest_mining_info else 0
            final_coord = latest_mining_info['坐标'] if latest_mining_info else ''
            final_shield = latest_mining_info['罩子'] if latest_mining_info else ''

        miners.append({
            '账号': account,
            '昵称': final_nickname,
            '联盟': final_alliance,
            '炉子': final_furnace,
            '炉子等级': final_furnace_level,
            '声望': final_prestige,
            '坐标': final_coord,
            '罩子': final_shield,
            '最大采集队数': max_teams,
            '挖矿天数': f"{mining_days}/{total_days}",
            '挖矿比例': round(ratio, 2),
            '备注': note,
            '_ratio': ratio,
            '_mining_days': mining_days,
            '_has_note': bool(note),
        })

df_miners = pd.DataFrame(miners)

if len(df_miners) > 0:
    df_miners = df_miners.sort_values(['_has_note', '_ratio', '_mining_days', '声望'],
                                      ascending=[False, False, False, False])
    df_miners = df_miners.drop(columns=['_ratio', '_mining_days'])
    df_miners = df_miners.reset_index(drop=True)
    df_miners.index = df_miners.index + 1
    df_miners.index.name = '序号'
    # 挖矿比例格式化
    df_miners['挖矿比例'] = df_miners['挖矿比例'].apply(lambda x: f'{x:.0%}')

    watched_count = df_miners['_has_note'].sum()
    print(f"\n✅ 高频挖矿小号: {len(df_miners)} 个 (其中标记账号: {watched_count} 个)")
else:
    print(f"\n❌ 没有找到满足条件的玩家")

# ==================== 保存（带样式） ====================
output_file = f'高频挖矿小号_{total_days}次分析.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    if len(df_miners) > 0:
        output_cols = ['账号', '昵称', '联盟', '炉子', '炉子等级', '声望',
                       '坐标', '罩子', '最大采集队数', '挖矿天数', '挖矿比例', '备注']
        df_miners[output_cols].to_excel(writer, sheet_name='高频挖矿小号', index=True)

        # 获取工作表并添加样式
        worksheet = writer.sheets['高频挖矿小号']

        # 设置红色加粗样式
        red_bold_font = Font(color='FF0000', bold=True, size=11)
        red_fill = PatternFill(start_color='FFE6E6', end_color='FFE6E6', fill_type='solid')

        # 遍历所有行，标记有备注的账号
        remark_count = 0
        for row_idx in range(2, len(df_miners) + 2):
            note_cell = worksheet.cell(row=row_idx, column=len(output_cols) + 1)  # +1因为有索引列
            note_value = note_cell.value

            if note_value and str(note_value).strip():
                remark_count += 1
                # 将整行设置为红色加粗
                for col_idx in range(1, len(output_cols) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.font = red_bold_font
                    cell.fill = red_fill

        # 设置表头样式
        header_font_white = Font(bold=True, size=11, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')

        for col_idx in range(1, len(output_cols) + 2):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = cell.alignment.copy(horizontal='center')

        # 自动调整列宽
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    cell_value = str(cell.value) if cell.value else ''
                    # 中文字符按2个字符宽度计算
                    char_length = 0
                    for char in cell_value:
                        if '\u4e00' <= char <= '\u9fff' or '\u3000' <= char <= '\u303f':
                            char_length += 2
                        else:
                            char_length += 1
                    max_length = max(max_length, char_length)
                except:
                    pass
            adjusted_width = min(max_length + 3, 35)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # 冻结首行
        worksheet.freeze_panes = 'A2'

        # 添加筛选器
        worksheet.auto_filter.ref = worksheet.dimensions

print(f"\n📁 结果已保存: {output_file}")

if len(df_miners) > 0:
    print(f"\n🔥 高频挖矿小号（前20）:")
    for i, (_, row) in enumerate(df_miners.head(20).iterrows()):
        note_mark = f" ⚠️[{row['备注']}]" if row['_has_note'] else ""
        print(
            f"  {i + 1}. [{row['联盟']}] {row['昵称']} | {row['炉子']} | 声望{row['声望']} | {row['最大采集队数']}队 | {row['挖矿天数']}{note_mark}")

# ==================== 生成分析摘要 ====================
print(f"\n📊 分析摘要:")
print(f"  - 分析时间跨度: {len(sorted_dates)}天")
print(f"  - 处理文件数量: {len(selected_files)}个")
print(f"  - 发现矿工数量: {len(df_miners)}个")
if len(df_miners) > 0:
    print(f"  - 标记账号数量: {watched_count}个")
    top_alliance = df_miners['联盟'].value_counts().head(1)
    if len(top_alliance) > 0:
        print(f"  - 最多矿工联盟: {top_alliance.index[0]} ({top_alliance.values[0]}个)")