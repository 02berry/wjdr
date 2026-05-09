"""聚落细节图 — 按聚落切块放大，标注玩家名"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
import re, os, warnings
import matplotlib.font_manager as fm
import matplotlib.colors as mcolors
import matplotlib.patheffects as pe
from matplotlib.cm import ScalarMappable
warnings.filterwarnings('ignore', message='Glyph')

for fp in ['C:/Windows/Fonts/simsun.ttc', 'C:/Windows/Fonts/cambria.ttc',
           'C:/Windows/Fonts/times.ttf', 'C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['SimSun', 'Times New Roman']
plt.rcParams['axes.unicode_minus'] = False

# ========== 配置 ==========
DATA_FOLDER = 'data'
MAP_FOLDER = 'Map'
MAPPING_FILE = f'{MAP_FOLDER}/league_mapping.xlsx'
S = 1200
PLAYER_SIZE = 2
DPI = 1600
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
BG_COLOR = '#f5f2ed'          # 暖白背景，降低与色块对比
other_color = '#ddd0bf'       # 非聚落成员色（比背景略深即可，不抢眼）

def _soften(c, amount=0.45):
    """颜色柔和化：向白色混合，减轻花眼"""
    v = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
    v = tuple(int(x + (255 - x) * amount) for x in v)
    return f'#{v[0]:02x}{v[1]:02x}{v[2]:02x}'

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

def find_bear_pit(member_idx, df, top_n=5):
    """在聚落中查找熊坑（3×3无玩家区域，靠近顶尖玩家）"""
    members = df.loc[member_idx]
    min_gx, max_gx = members['gx'].min(), members['gx'].max()
    min_gy, max_gy = members['gy'].min(), members['gy'].max()
    # 用区域内所有玩家（含其他联盟）检测空位，避免熊坑压到人
    in_area = df[(df['gx'] >= min_gx) & (df['gx'] <= max_gx) &
                  (df['gy'] >= min_gy) & (df['gy'] <= max_gy)]
    # 玩家占 (gx,gy)~(gx+1,gy+1) 四格，坐标指向底部格子
    occupied = set()
    for _, row in in_area.iterrows():
        for dx in (0, 1):
            for dy in (0, 1):
                occupied.add((row['gx'] + dx, row['gy'] + dy))
    top_centers = list(zip(
        members.nlargest(top_n, '声望')['gx'],
        members.nlargest(top_n, '声望')['gy']))
    sz = 3
    best, best_dist = None, float('inf')
    for x in range(min_gx, max_gx - (sz - 1)):
        for y in range(min_gy, max_gy - (sz - 1)):
            if any((x + dx, y + dy) in occupied for dx in range(sz) for dy in range(sz)):
                continue
            # 到顶尖玩家的平均距离（曼哈顿距离，按格算）
            avg = sum(abs(x - tx) + abs(y - ty) for tx, ty in top_centers) / top_n
            if avg < best_dist:
                best_dist, best = avg, (x, y)
    return best

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

# ========== 联盟基础色（每人深浅在渲染时按个人声望算） ==========
for s in settlements:
    alliance = s['name']
    if alliance in color_map:
        clr = color_map[alliance]
    else:
        while auto_idx < len(AUTO_PALETTE) and AUTO_PALETTE[auto_idx] in used_colors:
            auto_idx += 1
        clr = AUTO_PALETTE[auto_idx % len(AUTO_PALETTE)]
        auto_idx += 1
        used_colors.add(clr)
    settle_color[alliance] = clr

# ========== 熊坑检测 ==========
for s in settlements:
    bp = find_bear_pit(s['member_idx'], df)
    s['bear_pit'] = bp
    if bp:
        print(f'  {s["name"]} 熊坑: ({bp[0]}, {bp[1]})')

# ========== 逐聚落输出 ==========
out_dir = f'{MAP_FOLDER}/{date_str}'
os.makedirs(out_dir, exist_ok=True)

month, day = int(date_str[:2]), int(date_str[2:])

for s in settlements:
    vp = s['viewport']
    name = s['name']
    # 文件名安全
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)

    fig, ax = plt.subplots(figsize=(10, 10))
    fig.patch.set_facecolor(BG_COLOR)
    ax.set_facecolor(BG_COLOR)
    fig.subplots_adjust(right=0.85)  # 给 colorbar 腾空间

    # 视野内所有玩家（聚落成员 + 路过者）
    in_view = df[(df['dx'] >= vp[0]) & (df['dx'] <= vp[2]) &
                 (df['dy'] >= vp[1]) & (df['dy'] <= vp[3])]

    member_set = set(s['member_idx'])
    verts, colors = [], []
    gold_patches = []   # 过亿：(菱形顶点, 金色)
    fire_labels = []    # 火晶：(dx, dy, 等级数字)
    fire_verts = []     # 火晶：四段线单独画（避免被相邻方块覆盖）

    for i, row in in_view.iterrows():
        if i in member_set:
            base = settle_color.get(row['联盟'], other_color)
            prestige = row['声望']
            is_gold = prestige >= 10_000

            # 全局声望深浅映射（3000万~1亿，低于3000全最浅，高于1亿金黄方块）
            if is_gold:
                clr = None  # 金黄色方块，稍后单独处理
            elif prestige <= 3000:
                clr = _soften(base, 0.90)   # 最浅档（几乎只剩淡淡一层色）
            else:
                norm = (prestige - 3000) / 7000   # [0,1] 归一化到 3000~10000
                clr = _soften(base, 0.90 - norm * 0.90)  # 0.90→0.00

            furnace = str(row.get('炉子', ''))
            is_fire = furnace.startswith('火')

            # 火晶等级标注（过亿玩家不标）
            if is_fire and not is_gold:
                m = re.search(r'火(\d+)', furnace)
                if m:
                    fire_labels.append((row['dx'], row['dy'], m.group(1)))

            # 菱形顶点（玩家占 (gx,gy)~(gx+2,gy+2) 四格）
            v = [
                to_display(row['gx'], row['gy']),
                to_display(row['gx'] + 2, row['gy']),
                to_display(row['gx'] + 2, row['gy'] + 2),
                to_display(row['gx'], row['gy'] + 2),
            ]

            if is_gold:
                # 过亿 → 金色菱形块（声望越高越亮）
                gold_norm = min((prestige - 10000) / 5000, 1.0)  # 10000→0, 15000+→1
                r = int(0xDA + (0xFF - 0xDA) * gold_norm)
                g = int(0xA5 + (0xD7 - 0xA5) * gold_norm)
                b = int(0x20 + (0x00 - 0x20) * gold_norm)
                gold_patches.append((v, f'#{r:02x}{g:02x}{b:02x}'))
            else:
                verts.append(v)
                colors.append(clr)

            if is_fire and not is_gold:
                fire_verts.append(v)
        else:
            # 非成员：灰方块
            verts.append([
                to_display(row['gx'], row['gy']),
                to_display(row['gx'] + 2, row['gy']),
                to_display(row['gx'] + 2, row['gy'] + 2),
                to_display(row['gx'], row['gy'] + 2),
            ])
            colors.append(other_color)

    # 主方块层（统一灰网格线 0.3）
    if verts:
        ax.add_collection(PolyCollection(verts, facecolors=colors,
                                         edgecolors='#cccccc',
                                         linewidths=0.3, zorder=2))

    # 火晶金边（更高 zorder，不会被邻方块灰线覆盖）
    for v in fire_verts:
        ax.add_patch(plt.Polygon(v, fill=False, edgecolor='#FFE033',
                                 linewidth=0.6, zorder=2.1))

    # 过亿金色方块（声望越高越亮，无边）
    for gv, gc in gold_patches:
        ax.add_patch(plt.Polygon(gv, facecolor=gc, edgecolor='none', zorder=2.2))

    # 火晶等级标注（姓名上方）
    for dx, dy, lv in fire_labels:
        ax.text(dx, dy + 3.0, lv, fontsize=NAME_FONTSIZE * 2.0,
                color='#FFF44F', ha='center', va='center',
                fontweight='bold', zorder=4)

    # 玩家名字
    for i, row in in_view.iterrows():
        nick = str(row['昵称'])
        if not nick or nick == 'nan':
            continue
        if i in member_set:
            nm_clr = 'black'
        else:
            nm_clr = '#aaaaaa'
        nm_fs = NAME_FONTSIZE
        ax.text(row['dx'], row['dy'] + 2, split_name(nick),
                fontsize=nm_fs, color=nm_clr,
                ha='center', va='center', zorder=3)

    ax.set_xlim(vp[0], vp[2])
    ax.set_ylim(vp[1], vp[3])
    ax.set_aspect('equal')
    ax.axis('off')

    # 标题：x月x日[联盟]xx人
    ax.text(0.5, 0.97, f'{month}月{day}日[{name}]{s["n"]}人',
            transform=ax.transAxes, fontsize=24, fontweight='bold',
            ha='center', va='top', color=settle_color.get(name, '#444444'))

    # 熊坑标记（3×3方块，坐标指向底部格子）
    if s.get('bear_pit'):
        bx, by = s['bear_pit']
        bp_verts = [
            to_display(bx, by),         # 底部
            to_display(bx + 3, by),     # 右
            to_display(bx + 3, by + 3), # 顶部
            to_display(bx, by + 3),     # 左
        ]
        ax.add_patch(plt.Polygon(bp_verts, facecolor='black', edgecolor='none', zorder=6))

    # 声望深浅图例
    leg_base = settle_color.get(name, '#888888')
    cmap = mcolors.LinearSegmentedColormap.from_list('depth',
        [_soften(leg_base, 0.90), leg_base], N=256)
    cbar = fig.colorbar(ScalarMappable(cmap=cmap,
                        norm=mcolors.Normalize(3000, 10000)),
                        cax=fig.add_axes([0.87, 0.15, 0.025, 0.7]),
                        ticks=range(3000, 10001, 1000))
    cbar.set_label('战力（万）', fontsize=8)
    cbar.ax.tick_params(labelsize=6)

    out = f'{out_dir}/{date_str}_{safe_name}.png'
    try:
        fig.savefig(out, dpi=DPI, bbox_inches='tight')
    except OSError:
        # 旧文件被锁，先删再写
        try:
            os.remove(out)
        except FileNotFoundError:
            pass
        fig.savefig(out, dpi=DPI, bbox_inches='tight')
    plt.close(fig)
    print(f'  已保存: {out}')
