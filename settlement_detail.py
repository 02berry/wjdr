"""聚落细节图 — 按聚落切块放大，标注玩家名"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
import re, os
import matplotlib.font_manager as fm

for fp in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False

# ========== 配置 ==========
DATA_FOLDER = 'data'
MAP_FOLDER = 'Map'
MAPPING_FILE = f'{MAP_FOLDER}/league_mapping.xlsx'
S = 1200
PLAYER_SIZE = 2
DPI = 1200
MIN_SETTLEMENT = 50       # 只输出超过此人数的大聚落
VIEWPORT_PADDING = 0.35   # 视口边框额外扩充比例
VIEWPORT_SCALE = 1.0      # 地图放大倍数
NAME_FONTSIZE = 5

# ========== 名字宽度检测 ==========
def display_width(s):
    w = 0
    for c in s:
        w += 2 if ord(c) > 127 else 1
    return w

def split_name(s, max_w=10):
    """显示宽度 > max_w 则拆两行，从中点附近切"""
    if display_width(s) <= max_w:
        return s
    half = display_width(s) // 2
    cum = 0
    for i, c in enumerate(s):
        cum += 2 if ord(c) > 127 else 1
        if cum >= half:
            return s[:i] + '\n' + s[i:]
    return s

# ========== 读取最新数据 ==========
files = sorted([f for f in os.listdir(DATA_FOLDER)
                if f.startswith('3957_') and f.endswith('.xlsx')])
latest_file = os.path.join(DATA_FOLDER, files[-1])
print(f'读取: {latest_file}')

date_tag = files[-1].replace('3957_', '').replace('.xlsx', '')
# 去字母后缀取纯日期
date_str = re.sub(r'[a-zA-Z]', '', date_tag)

df = pd.read_excel(latest_file)

# ========== 解析坐标 ==========
coords = []
for c in df['坐标']:
    m = re.search(r'\((\d+),(\d+)\)', str(c))
    coords.append((int(m.group(1)), int(m.group(2))) if m else (None, None))
df['gx'], df['gy'] = zip(*coords)
df = df.dropna(subset=['gx', 'gy'])
df[['gx', 'gy']] = df[['gx', 'gy']].astype(int)

# ========== 联盟颜色 ==========
lm = pd.read_excel(MAPPING_FILE)
color_map = {}
for _, row in lm.iterrows():
    a, c = row['联盟'], row['颜色']
    if pd.notna(a) and pd.notna(c):
        color_map[str(a).strip()] = c
other_color = '#f5f0e0'
df['联盟'] = df['联盟'].astype(str).str.strip()

AUTO_PALETTE = [
    '#FF8C00', '#9370DB', '#20B2AA', '#D4A017', '#C71585',
    '#4169E1', '#D2691E', '#BA55D3', '#00CED1', '#FF69B4',
    '#8B4513', '#7B68EE', '#00BFFF', '#DB7093', '#6495ED',
    '#CD853F', '#DDA0DD', '#FF7F50', '#4682B4', '#F4A460',
]
used_colors = set(color_map.values())
auto_idx = 0
settle_color = {}

# ========== 坐标转换 ==========
def to_display(gx, gy):
    return gx - gy, gx + gy

df['dx'] = df['gx'] - df['gy']
df['dy'] = df['gx'] + df['gy']

# ========== 聚落检测 ==========
def find_settlement_pool(coords_list, min_size, max_radius=20):
    orig_idx = list(range(len(coords_list)))
    arr = np.array(coords_list, dtype=float)
    while len(arr) >= min_size:
        center = arr.mean(axis=0)
        dists = np.linalg.norm(arr - center, axis=1)
        if dists.max() < max_radius:
            return tuple(center), [orig_idx[i] for i in range(len(arr))]
        far = np.argmax(dists)
        arr = np.delete(arr, far, axis=0)
        del orig_idx[far]
    return None, []

# ========== 检测聚落并筛选 >=50 人 ==========
alliance_groups = df.groupby('联盟', sort=False)
alliance_power = alliance_groups['声望'].sum().sort_values(ascending=False)

settlements = []
for alliance, total_power in alliance_power.items():
    mask = df['联盟'] == alliance
    indices = df.index[mask].tolist()
    coords_list = [(df.at[i, 'gx'], df.at[i, 'gy']) for i in indices]
    center, kept = find_settlement_pool(coords_list, MIN_SETTLEMENT)
    if kept is None or len(kept) < MIN_SETTLEMENT:
        continue

    # 聚落颜色
    if alliance in color_map:
        clr = color_map[alliance]
    else:
        while auto_idx < len(AUTO_PALETTE) and AUTO_PALETTE[auto_idx] in used_colors:
            auto_idx += 1
        clr = AUTO_PALETTE[auto_idx % len(AUTO_PALETTE)]
        auto_idx += 1
        used_colors.add(clr)
    settle_color[alliance] = clr

    # 聚落成员原始索引
    member_idx = [indices[i] for i in kept]
    member_df = df.loc[member_idx]
    member_disp = member_df[['dx', 'dy']].values

    # 计算显示坐标包围盒
    cx, cy = member_disp.mean(axis=0)
    span = np.ptp(member_disp, axis=0).max()  # 取较长的边
    half = span / 2 * (1 + VIEWPORT_PADDING)
    half = half * VIEWPORT_SCALE

    vp_x0, vp_x1 = cx - half, cx + half
    vp_y0, vp_y1 = cy - half, cy + half

    settlements.append({
        'name': alliance,
        'power': int(total_power),
        'n': len(kept),
        'member_idx': member_idx,
        'center': (cx, cy),
        'viewport': (vp_x0, vp_y0, vp_x1, vp_y1),
    })
    print(f'  {alliance}: {len(kept)}人 (声望{int(total_power)}) → 视口{half:.0f}x{half:.0f}')

print(f'\n共 {len(settlements)} 个大聚落（≥{MIN_SETTLEMENT}人）')

# ========== 逐聚落输出 ==========
out_dir = f'{MAP_FOLDER}/聚落/{date_str}'
os.makedirs(out_dir, exist_ok=True)

half_p = PLAYER_SIZE / 2

for s in settlements:
    vp = s['viewport']
    name = s['name']
    # 文件名安全
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)

    fig, ax = plt.subplots(figsize=(10, 10))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    # 视野内所有玩家（聚落成员 + 路过者）
    in_view = df[(df['dx'] >= vp[0]) & (df['dx'] <= vp[2]) &
                 (df['dy'] >= vp[1]) & (df['dy'] <= vp[3])]

    member_set = set(s['member_idx'])
    verts, colors = [], []
    for i, row in in_view.iterrows():
        if i in member_set:
            clr = settle_color.get(row['联盟'], other_color)
        else:
            clr = other_color
        verts.append([
            to_display(row['gx'] - half_p, row['gy'] - half_p),
            to_display(row['gx'] + half_p, row['gy'] - half_p),
            to_display(row['gx'] + half_p, row['gy'] + half_p),
            to_display(row['gx'] - half_p, row['gy'] + half_p),
        ])
        colors.append(clr)

    if verts:
        ax.add_collection(PolyCollection(verts, facecolors=colors,
                                         edgecolors='#cccccc', linewidths=0.3, zorder=2))

    # 玩家名字
    for i, row in in_view.iterrows():
        nick = str(row['昵称'])
        if not nick or nick == 'nan':
            continue
        if i in member_set:
            nm_clr = 'black'
            nm_fs = NAME_FONTSIZE
        else:
            nm_clr = '#aaaaaa'
            nm_fs = NAME_FONTSIZE * 0.7
        ax.text(row['dx'], row['dy'], split_name(nick),
                fontsize=nm_fs, color=nm_clr,
                ha='center', va='center', zorder=3)

    ax.set_xlim(vp[0], vp[2])
    ax.set_ylim(vp[1], vp[3])
    ax.set_aspect('equal')
    ax.axis('off')

    out = f'{out_dir}/{safe_name}_{s["n"]}人.png'
    fig.savefig(out, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    print(f'  已保存: {out}')
