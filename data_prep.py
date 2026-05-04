"""
data_prep.py — 数据整理工具
1. 文件命名标准化：3957_MMDD.xlsx → 3957_MMDDa.xlsx
2. 声望修复：检测战力>10000进位bug，用前后文件插值修复
"""

import pandas as pd
import os
import glob
import re
from datetime import datetime

# ==================== 配置区 ====================
DATA_FOLDER = 'data'
FILE_PATTERN = '3957_*.xlsx'
PRESTIGE_HIGH = 10000     # 声望高于此值视为高战号
PRESTIGE_BUG_MAX = 100    # 声望低于此值且该号其他文件>10000则视为bug
DRY_RUN = False           # True=仅预览不实际操作

# ==================== 工具函数 ====================
def extract_date_parts(filename):
    """从文件名提取日期和序号"""
    match = re.search(r'3957_(\d{4})([a-z]?)\.xlsx', filename)
    if not match:
        return None, None, ''
    date_str = match.group(1)
    suffix = match.group(2)
    try:
        month, day = int(date_str[:2]), int(date_str[2:])
        date_obj = datetime(datetime.now().year, month, day)
        return date_obj, date_str, suffix
    except ValueError:
        return None, None, ''


def get_data_files():
    """获取排序后的数据文件列表"""
    files = glob.glob(os.path.join(DATA_FOLDER, FILE_PATTERN))
    files = [f for f in files if 'miner_remark' not in os.path.basename(f).lower()]
    files = [f for f in files if 'backup' not in f]

    dated_files = []
    for filepath in files:
        filename = os.path.basename(filepath)
        date_obj, date_str, suffix = extract_date_parts(filename)
        if date_obj:
            suffix_order = ord(suffix) - ord('a') + 1 if suffix else 0
            dated_files.append((filepath, date_obj, date_str, suffix, suffix_order))
        else:
            print(f"  ⚠️ 无法解析文件名: {filename}")

    dated_files.sort(key=lambda x: (x[1], x[4]))
    return dated_files


# ==================== 功能1：文件命名标准化 ====================
def normalize_filenames():
    """无后缀文件加a后缀，有b/c但无a的补a"""
    files = glob.glob(os.path.join(DATA_FOLDER, FILE_PATTERN))
    files = [f for f in files if 'miner_remark' not in os.path.basename(f).lower()]
    files = [f for f in files if 'backup' not in f]

    # 按日期统计有哪些后缀
    from collections import defaultdict
    day_files = defaultdict(list)
    for filepath in files:
        filename = os.path.basename(filepath)
        date_obj, date_str, suffix = extract_date_parts(filename)
        if date_obj:
            day_files[date_str].append((filepath, suffix))

    renamed = []
    for date_str, file_list in sorted(day_files.items()):
        suffixes = {s for _, s in file_list}
        dirname = os.path.dirname(file_list[0][0])

        # 情况1：有无后缀文件，且有b/c但没有a → 无后缀重命名为a
        # 情况2：只有无后缀文件（单文件日）→ 重命名为a
        for filepath, suffix in file_list:
            if suffix == '' and 'a' not in suffixes:
                new_name = f'3957_{date_str}a.xlsx'
                new_path = os.path.join(dirname, new_name)

                if os.path.exists(new_path):
                    print(f'  ⚠️ {os.path.basename(filepath)} → {new_name} 目标已存在，跳过')
                    continue

                if DRY_RUN:
                    print(f'  [预览] {os.path.basename(filepath)} → {new_name}')
                else:
                    os.rename(filepath, new_path)
                    print(f'  ✅ {os.path.basename(filepath)} → {new_name}')
                renamed.append((filepath, new_path))

    if not renamed:
        print('  ℹ️ 没有需要重命名的文件')
    return renamed


# ==================== 功能2：声望修复 ====================
def fix_prestige():
    """检测并修复战力进位bug"""
    dated_files = get_data_files()
    if not dated_files:
        print('  ❌ 没有数据文件')
        return

    # 1. 读取所有文件
    file_dfs = {}  # {filepath: DataFrame}
    for filepath, _, _, _, _ in dated_files:
        df = pd.read_excel(filepath, dtype=str)
        df['账号'] = pd.to_numeric(df['账号'], errors='coerce')
        df['_prestige_num'] = pd.to_numeric(df['声望'], errors='coerce')
        file_dfs[filepath] = df

    # 2. 构建声望时间线
    account_history = {}  # {account: [(filepath, file_idx, prestige), ...]}
    file_paths = [f[0] for f in dated_files]

    for file_idx, filepath in enumerate(file_paths):
        df = file_dfs[filepath]
        for _, row in df.iterrows():
            acc = int(row['账号']) if pd.notna(row['账号']) else 0
            if acc == 0:
                continue
            prestige = row['_prestige_num'] if pd.notna(row['_prestige_num']) else 0
            if acc not in account_history:
                account_history[acc] = []
            account_history[acc].append((filepath, file_idx, prestige))

    # 3. 检测bug并计算修复值
    fixes = []  # [(filepath, account, row_idx_in_file, old_val, new_val)]

    for acc, history in account_history.items():
        if len(history) < 2:
            continue

        max_prestige = max(h[2] for h in history)
        if max_prestige < PRESTIGE_HIGH:
            continue

        # 正常值列表
        normal = [(h[1], h[2]) for h in history if h[2] >= PRESTIGE_BUG_MAX]

        for filepath, file_idx, prestige in history:
            if prestige < PRESTIGE_BUG_MAX and prestige < max_prestige * 0.1:
                prev_val = None
                next_val = None
                for f_idx, val in normal:
                    if f_idx < file_idx:
                        prev_val = val
                    elif f_idx > file_idx and next_val is None:
                        next_val = val

                if prev_val is None and next_val is None:
                    continue
                elif prev_val is None:
                    fixed = next_val
                elif next_val is None:
                    fixed = prev_val
                else:
                    fixed = round((prev_val + next_val) / 2)

                # 找到在DataFrame中的行索引
                df = file_dfs[filepath]
                mask = df['账号'] == acc
                if mask.any():
                    row_idx = df[mask].index[0]
                    fixes.append((filepath, acc, row_idx, int(prestige), fixed))

    if not fixes:
        print('  ℹ️ 没有检测到声望异常')
        return

    print(f'\n  检测到 {len(fixes)} 个声望异常:')

    # 按文件分组
    from collections import defaultdict
    by_file = defaultdict(list)
    for f in fixes:
        by_file[f[0]].append(f)

    for filepath, file_fixes in by_file.items():
        filename = os.path.basename(filepath)
        print(f'\n  📄 {filename}:')
        for _, acc, _, old_val, new_val in file_fixes:
            print(f'    账号{acc}: {old_val} → {new_val}')

        if not DRY_RUN:
            # 修改DataFrame
            df = file_dfs[filepath]
            for _, acc, row_idx, _, new_val in file_fixes:
                df.at[row_idx, '声望'] = str(new_val)

            # 写回
            df.drop(columns=['_prestige_num'], inplace=True, errors='ignore')
            df.to_excel(filepath, index=False)
            print(f'  ✅ 已保存')

    if DRY_RUN:
        print(f'\n  ℹ️ 预览模式，未实际修改。共{len(fixes)}处待修复。')


# ==================== 功能3：联盟列文本化 ====================
def normalize_alliance():
    """将所有文件的联盟列转为纯文本，防止数字/文本混用导致误报变化"""
    dated_files = get_data_files()
    if not dated_files:
        return

    fixed = 0
    for filepath, _, _, _, _ in dated_files:
        df = pd.read_excel(filepath, dtype=str)
        if '联盟' not in df.columns:
            continue

        changed = False
        for idx, row in df.iterrows():
            raw = row['联盟']
            if pd.notna(raw):
                cleaned = str(raw).strip()
                if cleaned != str(raw):
                    df.at[idx, '联盟'] = cleaned
                    changed = True
                # 确保是文本格式（非纯数字字符串也保持）
            elif pd.isna(raw):
                df.at[idx, '联盟'] = '-'
                changed = True

        if changed:
            if not DRY_RUN:
                df.to_excel(filepath, index=False)
            fixed += 1
            print(f'  ✅ {os.path.basename(filepath)} 联盟列已文本化')

    if fixed == 0:
        print('  ℹ️ 联盟列格式正常，无需修复')


# ==================== 主流程 ====================
if __name__ == '__main__':
    mode = '🔍预览' if DRY_RUN else '🔧执行'
    print(f'{"=" * 50}')
    print(f'📋 数据整理工具 [{mode}]')
    print(f'{"=" * 50}')

    print('\n1️⃣ 文件命名标准化...')
    normalize_filenames()

    print('\n2️⃣ 修复声望进位bug...')
    fix_prestige()

    print('\n3️⃣ 联盟列文本化...')
    normalize_alliance()

    print(f'\n✅ 完成')
