"""游戏地图可视化 — 菱形地图 + 地形分区 + 聚落识别"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PolyCollection
import re, os
from scipy import stats
import matplotlib.font_manager as fm

# 显式注册中文字体，避免缓存导致的回退到 DejaVu Sans
for fp in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
# Nature-journal style
plt.rcParams.update({
    'axes.spines.right': False,
    'axes.spines.top': False,
    'axes.linewidth': 0.8,
    'legend.frameon': False,
    'font.size': 7,
    'svg.fonttype': 'none',
})

# ========== 配置 ==========
DATA_FOLDER = 'data'
MAP_FOLDER = 'Map'
MAPPING_FILE = f'{MAP_FOLDER}/league_mapping.xlsx'
ARCH_FILE = f'{MAP_FOLDER}/Architecture.xlsx'
S = 1200
PLAYER_SIZE = 2
DPI = 1200

# 矿工高亮
MINER_FILE = '21_miner.xlsx'
MINER_COLOR = '#000000'       # 纯黑，非常显眼

# 聚落标注手动偏移（联盟名 → (dx, dy) 显示坐标偏移量）
LABEL_OFFSET = {
    'DXY': (30, 10),
}
SETTLEMENT_MIN = 20     # 聚落最小人数
MAX_RADIUS = 30         # 聚落半径上限

# 自动配色（非红非绿，互相好区分）
AUTO_PALETTE = [
    '#FF8C00', '#9370DB', '#20B2AA', '#D4A017', '#C71585',
    '#4169E1', '#D2691E', '#BA55D3', '#00CED1', '#FF69B4',
    '#8B4513', '#7B68EE', '#00BFFF', '#DB7093', '#6495ED',
    '#CD853F', '#DDA0DD', '#FF7F50', '#4682B4', '#F4A460',
]

# 地形配置
ZONES = [
    ('荒原', 1200, '#fff8e1', '#ffe082'),
    ('雪原',  500, '#e3f2fd', '#90caf9'),
    ('沃土',  200, '#e8f5e9', '#a5d6a7'),
]

# ========== 读取最新数据 ==========
files = sorted([f for f in os.listdir(DATA_FOLDER)
                if f.startswith('3957_') and f.endswith('.xlsx')])
latest_file = os.path.join(DATA_FOLDER, files[-1])
print(f'读取: {latest_file}')

df = pd.read_excel(latest_file)

# ========== 解析坐标 ==========
coords = []
for c in df['坐标']:
    m = re.search(r'\((\d+),(\d+)\)', str(c))
    coords.append((int(m.group(1)), int(m.group(2))) if m else (None, None))

df['gx'], df['gy'] = zip(*coords)
df = df.dropna(subset=['gx', 'gy'])
df[['gx', 'gy']] = df[['gx', 'gy']].astype(int)

# ========== 联盟颜色映射表 ==========
lm = pd.read_excel(MAPPING_FILE)
color_map = {}
for _, row in lm.iterrows():
    a, c = row['联盟'], row['颜色']
    if pd.notna(a) and pd.notna(c):
        color_map[str(a).strip()] = c

other_color = '#f5f0e0'  # 接近背景色，缩小后几乎不可见
df['联盟'] = df['联盟'].astype(str).str.strip()

# ========== 读取矿工名单 ==========
miner_df = pd.read_excel(MINER_FILE)
miner_accounts = set(miner_df['账号'].tolist())
df['is_miner'] = df['账号'].isin(miner_accounts)
print(f'矿工玩家: {df["is_miner"].sum()}')

# ========== 读取建筑 ==========
arch = pd.read_excel(ARCH_FILE)
sun = arch[arch['名称'] == '太阳城'].iloc[0]
sun_cx = int(str(sun['坐标'])[:3])
sun_cy = int(str(sun['坐标'])[3:])
print(f'太阳城: ({sun_cx}, {sun_cy})')

# ========== 坐标转换（菱形显示） ==========
def to_display(gx, gy):
    return gx - gy, gx + gy

def zone_diamond(cx, cy, half):
    corners = [
        (cx - half, cy - half), (cx + half, cy - half),
        (cx + half, cy + half), (cx - half, cy + half),
    ]
    return [to_display(x, y) for x, y in corners]

# 预计算显示坐标
df['dx'] = df['gx'] - df['gy']
df['dy'] = df['gx'] + df['gy']

# ========== 聚落检测 ==========
def find_settlement_pool(coords_list, min_size, max_radius):
    """反复去掉离中心最远的玩家，直到全部在 max_radius 内。"""
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


def find_label_pos(center_dx, center_dy, all_players, label_w=48, label_h=26,
                   manual_offset=None):
    """从中心外扩，优先0遮挡，其次≤1遮挡。manual_offset=(dx, dy) 强制偏移。"""
    if manual_offset is not None:
        return center_dx + manual_offset[0], center_dy + manual_offset[1]

    pts = np.array(all_players)
    if len(pts) == 0:
        return center_dx, center_dy + 20

    best_pos_0 = None   # 0 遮挡
    best_pos_1 = None   # ≤1 遮挡
    best_dist_0 = float('inf')
    best_dist_1 = float('inf')

    for angle_deg in range(0, 360, 15):
        rad = np.radians(angle_deg)
        dx, dy = np.cos(rad), np.sin(rad)

        for dist in range(0, 100, 3):
            cx = center_dx + dist * dx
            cy_bottom = center_dy + dist * dy

            overlap_cnt = 0
            for px, py in pts:
                if (abs(px - cx) < label_w / 2 and
                    cy_bottom <= py <= cy_bottom + label_h):
                    overlap_cnt += 1
                    if overlap_cnt > 1:
                        break

            if overlap_cnt == 0:
                if dist < best_dist_0:
                    best_dist_0 = dist
                    best_pos_0 = np.array([cx, cy_bottom])
                break  # 该方向找到0遮挡，不用继续外扩
            elif overlap_cnt <= 1 and best_pos_0 is None:
                if dist < best_dist_1:
                    best_dist_1 = dist
                    best_pos_1 = np.array([cx, cy_bottom])

    if best_pos_0 is not None:
        return best_pos_0[0], best_pos_0[1]
    if best_pos_1 is not None:
        return best_pos_1[0], best_pos_1[1]
    return center_dx, center_dy + 30

# ========== 逐联盟检测聚落 ==========
alliance_groups = df.groupby('联盟', sort=False)
alliance_power = alliance_groups['声望'].sum().sort_values(ascending=False)

df['color'] = None
df['in_settlement'] = False

used_colors = set(color_map.values())
auto_idx = 0
settlement_alliances = {}
settlement_labels = []

for alliance, total_power in alliance_power.items():
    mask = df['联盟'] == alliance
    indices = df.index[mask].tolist()
    coords_list = [(df.at[i, 'gx'], df.at[i, 'gy']) for i in indices]

    center, kept = find_settlement_pool(coords_list, SETTLEMENT_MIN, MAX_RADIUS)
    if kept is None or len(kept) < SETTLEMENT_MIN:
        continue

    if alliance in color_map:
        clr = color_map[alliance]
    else:
        while auto_idx < len(AUTO_PALETTE) and AUTO_PALETTE[auto_idx] in used_colors:
            auto_idx += 1
        clr = AUTO_PALETTE[auto_idx % len(AUTO_PALETTE)]
        auto_idx += 1
        used_colors.add(clr)

    original_idx = [indices[i] for i in kept]
    for i in original_idx:
        df.at[i, 'color'] = clr
        df.at[i, 'in_settlement'] = True

    settlement_alliances[alliance] = clr

    # 计算标注位置：360度扫描，对所有玩家做遮挡检测
    member_disp = [to_display(df.at[i, 'gx'], df.at[i, 'gy']) for i in original_idx]
    all_disp = list(zip(df['dx'], df['dy']))     # 所有玩家
    cdx = center[0] - center[1]
    cdy = center[0] + center[1]
    lx, ly = find_label_pos(cdx, cdy, all_disp,
                              manual_offset=LABEL_OFFSET.get(alliance))
    settlement_labels.append((lx, ly, alliance, clr))
    print(f'  {alliance} (声望{int(total_power)}): {len(kept)}人 → {clr}')

# ========== 非聚落玩家着色 ==========
no_settle = df[~df['in_settlement']].index
for i in no_settle:
    a = df.at[i, '联盟']
    df.at[i, 'color'] = color_map.get(a, other_color)

n_settlement = df['in_settlement'].sum()
print(f'聚落总数: {n_settlement}')

# ========== 矿工高亮（保存原色后覆盖） ==========
df['pre_miner_color'] = df['color'].copy()
df.loc[df['is_miner'], 'color'] = MINER_COLOR
n_miner = df['is_miner'].sum()
print(f'矿工高亮: {n_miner}人 → {MINER_COLOR}')

# ========== 绘图 ==========
fig = plt.figure(figsize=(12, 13))
ax = fig.add_axes([-0.02, 0.14, 1.04, 0.78])  # 上下留边给副图
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# ========== 地形分区 ==========
for name, side, fill_clr, line_clr in ZONES:
    half = side // 2
    verts = zone_diamond(sun_cx, sun_cy, half)
    poly = Polygon(verts, fill=True, facecolor=fill_clr,
                   edgecolor=line_clr, linewidth=1.5, linestyle='--', zorder=0)
    ax.add_patch(poly)

# ========== 玩家（普通 + 双色矿工） ==========
half_p = PLAYER_SIZE / 2
normal_verts, normal_colors = [], []
dual_left_verts, dual_left_colors = [], []
dual_right_verts = []
has_main = df['pre_miner_color'] != other_color

for _, row in df.iterrows():
    gx, gy = row['gx'], row['gy']
    clr = row['color']
    is_miner, main_clr = row['is_miner'], has_main.loc[row.name]
    orig_clr = row['pre_miner_color']

    if is_miner and main_clr:
        # 竖切：左半主色，右半黑（用原始四角顶点切分）
        dual_left_verts.append([
            to_display(gx - 1, gy + 1),   # TL / 菱形左顶点
            to_display(gx - 1, gy - 1),   # BL / 菱形上顶点
            to_display(gx, gy),            # 中心
            to_display(gx + 1, gy + 1),   # TR / 菱形下顶点
        ])
        dual_left_colors.append(orig_clr)
        dual_right_verts.append([
            to_display(gx, gy),            # 中心
            to_display(gx - 1, gy - 1),   # BL / 菱形上顶点
            to_display(gx + 1, gy - 1),   # BR / 菱形右顶点
            to_display(gx + 1, gy + 1),   # TR / 菱形下顶点
        ])
    else:
        normal_verts.append([
            to_display(gx - half_p, gy - half_p),
            to_display(gx + half_p, gy - half_p),
            to_display(gx + half_p, gy + half_p),
            to_display(gx - half_p, gy + half_p),
        ])
        normal_colors.append(MINER_COLOR if is_miner else clr)

if normal_verts:
    ax.add_collection(PolyCollection(normal_verts, facecolors=normal_colors,
                                     edgecolors='none', zorder=2))
if dual_left_verts:
    ax.add_collection(PolyCollection(dual_left_verts, facecolors=dual_left_colors,
                                     edgecolors='none', zorder=2))
    ax.add_collection(PolyCollection(dual_right_verts,
                                     facecolors=[MINER_COLOR] * len(dual_right_verts),
                                     edgecolors='none', zorder=2))

# ========== 聚落标注 ==========
for dx, dy, name, clr in settlement_labels:
    ax.text(dx, dy, name, ha='center', va='bottom', fontsize=9,
            color=clr, fontweight='bold', zorder=5)

# ========== 建筑 ==========
for _, b in arch.iterrows():
    name, coord_str, side = b['名称'], str(b['坐标']), int(b['边长'])
    cx, cy = int(coord_str[:3]), int(coord_str[3:])
    half = side / 2
    corners = [
        (cx - half, cy - half), (cx + half, cy - half),
        (cx + half, cy + half), (cx - half, cy + half),
    ]
    verts = [to_display(x, y) for x, y in corners]
    poly = Polygon(verts, fill=True, facecolor='#e8dcc8',
                   edgecolor='#8b7355', linewidth=1.5, zorder=3)
    ax.add_patch(poly)
    dx, dy = to_display(cx, cy)
    fs, fw = (9, 'bold') if name == '太阳城' else (7, 'normal')
    ax.text(dx, dy + 8, name, ha='center', va='bottom', fontsize=fs,
            color='#5c4033', fontweight=fw, zorder=4)

# ========== 图例 ==========
legend_handles = []
for a, c in color_map.items():
    if a == 'other':
        continue
    cnt = len(df[df['联盟'] == a])
    if cnt > 0:
        legend_handles.append(plt.Line2D([0], [0], marker='s', color='w',
                               markerfacecolor=c, markersize=8, label=f'{a} ({cnt})'))
for a, c in settlement_alliances.items():
    if a in color_map:
        continue
    cnt = len(df[(df['联盟'] == a) & df['in_settlement']])
    if cnt > 0:
        legend_handles.append(plt.Line2D([0], [0], marker='s', color='w',
                               markerfacecolor=c, markersize=8, label=f'{a} ({cnt})'))
other_cnt = len(df[~df['联盟'].isin(color_map) & ~df['in_settlement']])
if other_cnt > 0:
    legend_handles.append(plt.Line2D([0], [0], marker='s', color='w',
                           markerfacecolor=other_color, markersize=8,
                           label=f'其他 ({other_cnt})'))

ax.legend(handles=legend_handles, loc='upper right', framealpha=0.9,
          facecolor='white', edgecolor='#cccccc', labelcolor='#333',
          fontsize=9, title='联盟', title_fontsize=10)

# ========== 地图边框 ==========
map_border = Polygon([
    to_display(0, 0), to_display(S, 0),
    to_display(S, S), to_display(0, S),
], fill=False, edgecolor='#999', linewidth=2, zorder=0)
ax.add_patch(map_border)

# ========== 标题 ==========
date_tag = files[-1].replace('3957_', '').replace('.xlsx', '')
date_short = date_tag[:4]
OUTPUT_FILE = f'{MAP_FOLDER}/地图_{date_short}.png'
month, day = date_tag[:2], date_tag[2:4]
ax.set_title(f'无尽冬日 地图分布 ({int(month)}月{int(day)}日)',
             color='#333', fontsize=16, pad=20)

# ========== 四角统计可视化 ==========

def style_ax(ch):
    """Nature风格：保留左/下轴，刻度朝内，无网格"""
    ch.spines['top'].set_visible(False)
    ch.spines['right'].set_visible(False)
    ch.spines['left'].set_linewidth(0.6)
    ch.spines['bottom'].set_linewidth(0.6)
    ch.tick_params(direction='in', length=3, width=0.5, labelsize=6.5)

# 左上：玩家人数时间序列
files_all = sorted([f for f in os.listdir(DATA_FOLDER)
                    if f.startswith('3957_') and f.endswith('.xlsx')])
# 按日期分组，取每日最后一份
date_groups = {}
for f in files_all:
    tag = f.replace('3957_', '').replace('.xlsx', '')
    date_key = tag[:4]  # MMDD
    date_groups[date_key] = f  # 后覆盖 = 取最后一份
daily_dates = sorted(date_groups.keys())
daily_counts = []
for d in daily_dates:
    f = date_groups[d]
    cdf = pd.read_excel(f'{DATA_FOLDER}/{f}')
    daily_counts.append(len(cdf))

ax_tl = fig.add_axes([0.04, 0.92, 0.20, 0.06])
ax_tl.plot(range(len(daily_counts)), daily_counts, 'o-', color='#0F4D92',
           markersize=2.5, linewidth=1.0)
ax_tl.fill_between(range(len(daily_counts)), daily_counts, alpha=0.12, color='#0F4D92')
ax_tl.set_xlim(-0.3, len(daily_counts) - 0.7)
# 只标首尾和最大日期
n = len(daily_counts)
ticks, labels = [], []
ticks.append(0); labels.append(f'{daily_dates[0][:2]}/{daily_dates[0][2:]}')
if n > 1:
    max_idx = daily_counts.index(max(daily_counts))
    if 0 < max_idx < n-1:
        ticks.append(max_idx); labels.append(f'{daily_dates[max_idx][:2]}/{daily_dates[max_idx][2:]}\n{max(daily_counts)}人')
    ticks.append(n-1); labels.append(f'{daily_dates[-1][:2]}/{daily_dates[-1][2:]}')
ax_tl.set_xticks(ticks)
ax_tl.set_xticklabels(labels, fontsize=6)
ax_tl.tick_params(axis='y', labelsize=6)
ax_tl.set_ylabel('人数', fontsize=7)
style_ax(ax_tl)

# --- 左下：炉子等级分布（逐级条形图，Nature风格）---
ax_bl = fig.add_axes([0.05, 0.02, 0.40, 0.09])
furnace_counts = {}
for v in df['炉子']:
    s = str(v)
    if s.startswith('火'):
        furnace_counts['火'] = furnace_counts.get('火', 0) + 1
    else:
        try:
            lv = int(s.replace('lv.', ''))
            furnace_counts[lv] = furnace_counts.get(lv, 0) + 1
        except:
            pass
levels = sorted([k for k in furnace_counts if isinstance(k, int)])
all_labels = [str(lv) for lv in levels] + ['火']
all_counts = [furnace_counts[lv] for lv in levels] + [furnace_counts.get('火', 0)]
xp = np.arange(len(all_labels))
ax_bl.bar(xp, all_counts, color='#3775BA', edgecolor='white', linewidth=0.3, width=0.7)
style_ax(ax_bl)
tick_step = max(1, len(all_labels) // 12)
ax_bl.set_xticks(xp[::tick_step])
ax_bl.set_xticklabels([all_labels[i] for i in range(0, len(all_labels), tick_step)], fontsize=5.5)
ax_bl.set_xlabel('炉子等级', fontsize=6.5)
ax_bl.set_ylabel('人数', fontsize=6.5)

# --- 右下：声望分布直方图 + 正态拟合（Nature风格）---
ax_br = fig.add_axes([0.55, 0.02, 0.42, 0.09])
pv = df['声望'].values / 10000
mu, std = stats.norm.fit(pv)
bin_edges = np.arange(0, df['声望'].max() + 500, 500) / 10000
counts, bins_out, patches = ax_br.hist(pv, bins=bin_edges, color='#0F4D92',
                                        edgecolor='white', linewidth=0.3, alpha=0.85)
style_ax(ax_br)
ax_br.set_xlabel('声望（万）', fontsize=6.5)
ax_br.set_ylabel('人数', fontsize=6.5)
xf = np.linspace(pv.min(), pv.max(), 200)
yf = stats.norm.pdf(xf, mu, std) * len(pv) * (bins_out[1] - bins_out[0])
ax_br.plot(xf, yf, linewidth=1.0, alpha=0.6, color='#B64342')
for bar, c in zip(patches, counts):
    if c > 0:
        ax_br.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                   f'{int(c)}', ha='center', va='bottom', fontsize=5)

ax.set_xlim(-S - 80, S + 80)
ax.set_ylim(-80, S * 2 + 80)
ax.set_aspect('equal')
ax.axis('off')

fig.savefig(OUTPUT_FILE, dpi=DPI, bbox_inches='tight')
fig.savefig(OUTPUT_FILE.replace('.png', '.svg'), bbox_inches='tight')
print(f'已保存: {OUTPUT_FILE} (DPI={DPI})')
print(f'已保存: {OUTPUT_FILE.replace(".png", ".svg")}')
print(f'玩家总数: {len(df)}')
