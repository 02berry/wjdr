"""全地图可视化 — 菱形地图 + 地形分区 + 聚落标注 + 旗子领地 + 熊坑

输出: Map/地图_MMDD.png (1200DPI, 12×12英寸, 全图约2400×2400显示单位)

坐标模型:
  to_display(gx, gy) = (gx - gy, gx + gy)
  这等价于把游戏坐标逆时针旋转45度并放大√2倍。
  玩家坐标 (gx,gy) 指向所占 2×2 块的底部格子。
  玩家菱形顶点: (gx,gy)→(gx+2,gy)→(gx+2,gy+2)→(gx,gy+2)
  地图几何中心: (600,600)，地形以同心菱形环排列。

渲染分层:
  zorder 0:   地形（荒原→雪原→沃土→废墟→灰烬→王域）
  zorder 1:   要塞/堡垒废墟方块（覆盖原本地形）
  zorder 1.5: 旗子领地填充（浅蓝）
  zorder 1.6: 领地联合边界（格子级4邻域，重叠处无线）
  zorder 2:   玩家菱形块（声望深浅着色）
  zorder 2.1: 火晶金边
  zorder 3:   建筑（太阳城、要塞、堡垒）
  zorder 5:   旗子圆 / HQ圆 / 聚落标注
  zorder 6:   熊坑图片

着色规则:
  前10联盟：联盟色 + 声望深浅（≤3000最浅，3000~10000渐变，>10000变深）
  非前10联盟：黑色调，纯黑对应全服最高声望，低声望依次减淡
  火晶玩家：金色边框线（0.15pt）

配置区（脚本顶部）:
  DATA_FOLDER, MAP_FOLDER, MAPPING_FILE, S=1200, DPI=1200
  SHOW_NAMES, LABEL_OFFSET, SETTLEMENT_MIN=20, MAX_RADIUS=30
  TOP_N=10, NEUTRAL='#000000'
  地形颜色见 ZONES 列表
  建筑坐标见 BUILDINGS 列表
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PolyCollection
import re, os
from scipy import stats
import matplotlib.font_manager as fm
from PIL import Image

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
S = 1200
PLAYER_SIZE = 2
DPI = 1200

# 玩家名称显示
SHOW_NAMES = False

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

# 地形配置（half = 中心到菱形顶点距离，所有地形同心于 MAP_CX,MAP_CY）
ZONES = [
    ('荒原', 599.5, '#fffcf0', '#ffe082'),
    ('雪原', 299.5, '#f1f9fd', '#90caf9'),
    ('沃土', 149.5, '#f4faf4', '#a5d6a7'),
    ('废墟',  48.5, '#f0f0f0', '#f0f0f0'),
    ('灰烬',  14, '#f2ecec', '#e0d6d6'),
]

# 建筑（坐标指向底部格子）
BUILDINGS = [
    ('太阳城', 597, 597, 6),
    ('1号要塞', 597, 800, 6),
    ('2号要塞', 400, 597, 6),
    ('3号要塞', 597, 400, 6),
    ('4号要塞', 800, 597, 6),
    ('1号堡垒', 237, 828, 6),
    ('2号堡垒', 237, 606, 6),
    ('3号堡垒', 237, 348, 6),
    ('4号堡垒', 366, 237, 6),
    ('5号堡垒', 588, 237, 6),
    ('6号堡垒', 846, 237, 6),
    ('7号堡垒', 957, 348, 6),
    ('8号堡垒', 957, 606, 6),
    ('9号堡垒', 957, 828, 6),
    ('10号堡垒', 828, 957, 6),
    ('11号堡垒', 606, 957, 6),
    ('12号堡垒', 348, 957, 6),
]

# ========== 读取建筑表（arc.xlsx）==========
ARC_FILE = f'{MAP_FOLDER}/arc.xlsx'
ARC_CFG = {
    'flag': {'size': 1, 'territory': 7, 'label': None},
    'Headquarters': {'size': 3, 'territory': 15, 'label': '联盟总部'},
}
if os.path.exists(ARC_FILE):
    rows = []
    xls = pd.ExcelFile(ARC_FILE)
    for sheet in xls.sheet_names:
        sdf = pd.read_excel(xls, sheet)
        col1, col2 = sdf.columns[0], sdf.columns[1]
        for _, r in sdf.iterrows():
            coord = r[col1]
            btype = 'flag'
            if pd.notna(r[col2]):
                val = str(r[col2]).strip()
                if val:
                    btype = val
            rows.append({'联盟': sheet, '类型': btype, '坐标': coord})
    arc_df = pd.DataFrame(rows)
    arc_df['gx'] = (arc_df['坐标'] // 1000).astype(int)
    arc_df['gy'] = (arc_df['坐标'] % 1000).astype(int)
    arc_df = arc_df.drop_duplicates(subset=['gx', 'gy'])
    arc_df['dx'] = arc_df['gx'] - arc_df['gy']
    arc_df['dy'] = arc_df['gx'] + arc_df['gy']
    print(f'建筑: {len(arc_df)} 座')

    # 计算所有领地格子（游戏坐标，用于联合边界）
    territory_cells = set()
    for _, b in arc_df.iterrows():
        cfg = ARC_CFG[b['类型']]
        half = cfg['territory']
        cx, cy = b['dx'], b['dy'] + cfg['size']
        fx, fy = b['gx'], b['gy']
        span = half + cfg['size']
        for gx in range(max(0, fx - span), min(S, fx + span + 1)):
            for gy in range(max(0, fy - span), min(S, fy + span + 1)):
                dx, dy = gx - gy, gx + gy
                if abs(dx - cx) + abs(dy + 1 - cy) <= half + 1e-9:
                    territory_cells.add((gx, gy))
    print(f'领地格子: {len(territory_cells)}')
else:
    arc_df = pd.DataFrame()
    territory_cells = set()
    print('建筑表不存在，跳过')

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

df['联盟'] = df['联盟'].astype(str).str.strip()

# 地图几何中心（用于地形分区）
MAP_CX, MAP_CY = 600, 600

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


def find_bear_pit(member_idx, df, top_n=5):
    """在聚落中查找熊坑（3×3无玩家区域，靠近顶尖玩家）"""
    members = df.loc[member_idx]
    min_gx, max_gx = members['gx'].min(), members['gx'].max()
    min_gy, max_gy = members['gy'].min(), members['gy'].max()
    in_area = df[(df['gx'] >= min_gx) & (df['gx'] <= max_gx) &
                  (df['gy'] >= min_gy) & (df['gy'] <= max_gy)]
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
            avg = sum(abs(x - tx) + abs(y - ty) for tx, ty in top_centers) / top_n
            if avg < best_dist:
                best_dist, best = avg, (x, y)
    return best


# ========== 前10联盟全联盟着色 ==========
alliance_power = df.groupby('联盟')['声望'].sum().sort_values(ascending=False)
# 排除无联盟标记
alliance_power = alliance_power[alliance_power.index != '-']
TOP_N = 10
top_alliances = set(alliance_power.head(TOP_N).index)
print(f'前{TOP_N}大联盟着色:')

used_colors = set(color_map.values())
auto_idx = 0
df['color'] = None

for alliance in alliance_power.index:
    if alliance not in top_alliances:
        continue
    if alliance in color_map:
        clr = color_map[alliance]
    else:
        while auto_idx < len(AUTO_PALETTE) and AUTO_PALETTE[auto_idx] in used_colors:
            auto_idx += 1
        clr = AUTO_PALETTE[auto_idx % len(AUTO_PALETTE)]
        auto_idx += 1
        used_colors.add(clr)
    df.loc[df['联盟'] == alliance, 'color'] = clr
    cnt = (df['联盟'] == alliance).sum()
    print(f'  {alliance}: {clr} ({cnt}人)')

# 其余玩家统一浅色
NEUTRAL = '#000000'
df['color'] = df['color'].fillna(NEUTRAL)

# ========== 聚落检测（标注位置 + 熊坑）==========
settlements = []
for alliance in alliance_power.head(TOP_N).index:
    mask = df['联盟'] == alliance
    indices = df.index[mask].tolist()
    coords_list = [(df.at[i, 'gx'], df.at[i, 'gy']) for i in indices]
    center, kept = find_settlement_pool(coords_list, SETTLEMENT_MIN, MAX_RADIUS)
    if kept is None or len(kept) < SETTLEMENT_MIN:
        continue
    clr = df.at[indices[0], 'color']
    all_disp = list(zip(df['dx'], df['dy']))
    cdx, cdy = center[0] - center[1], center[0] + center[1]
    lx, ly = find_label_pos(cdx, cdy, all_disp,
                              manual_offset=LABEL_OFFSET.get(alliance))
    member_idx = [indices[i] for i in kept]
    bp = find_bear_pit(member_idx, df)
    settlements.append({
        'name': alliance, 'color': clr,
        'label': (lx, ly), 'bear_pit': bp,
    })
    print(f'  {alliance} 标注: ({lx:.0f}, {ly:.0f}){" 熊坑" if bp else ""}')

# ========== 绘图 ==========
fig = plt.figure(figsize=(12, 12))
ax = fig.add_axes([0.07, 0.07, 0.86, 0.86])  # 外切正方形，居中
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# ========== 地形分区 ==========
for name, half, fill_clr, line_clr in ZONES:
    verts = zone_diamond(MAP_CX, MAP_CY, half)
    poly = Polygon(verts, fill=True, facecolor=fill_clr,
                   edgecolor='none', zorder=0)
    ax.add_patch(poly)

# ========== 王域（太阳城专属地形，中心 (600,600)）==========
KINGDOM_CX, KINGDOM_CY = 600, 600
KINGDOM_HALF = 6.0
verts = zone_diamond(KINGDOM_CX, KINGDOM_CY, KINGDOM_HALF)
ax.add_patch(Polygon(verts, fill=True, facecolor='#fff3cd',
                     edgecolor='none', zorder=0))

# 炮台（王域四角，2×2方块+内切圆）
TURRET_BASE = '#ffe082'
TURRET_CIRCLE = '#ffb300'
TURRET_R = 1.4
for gx, gy in [(593,593), (605,593), (605,605), (593,605)]:
    v = [to_display(gx, gy), to_display(gx+2, gy),
         to_display(gx+2, gy+2), to_display(gx, gy+2)]
    ax.add_patch(Polygon(v, fill=True, facecolor=TURRET_BASE,
                         edgecolor='none', zorder=2.5))
    cx, cy = to_display(gx+1, gy+1)
    ax.add_patch(plt.Circle((cx, cy), radius=TURRET_R,
                            facecolor=TURRET_CIRCLE, edgecolor='none', zorder=2.6))

# ========== 要塞/堡垒废墟地形（覆盖原本雪原/沃土）==========
RUINS_HALF = 29.5
RUINS_COLOR = '#f0f0f0'   # 和废墟地形同色
for name, gx, gy, side in BUILDINGS:
    if '要塞' not in name and '堡垒' not in name:
        continue
    cx, cy = gx + side / 2, gy + side / 2
    verts = zone_diamond(cx, cy, RUINS_HALF)
    poly = Polygon(verts, fill=True, facecolor=RUINS_COLOR,
                   edgecolor='none', zorder=1)
    ax.add_patch(poly)

# ========== 旗子领地（联合边界，重叠部分无线）==========
if len(arc_df) > 0:
    for _, b in arc_df.iterrows():
        fdx, fdy = b['dx'], b['dy']
        cfg = ARC_CFG[b['类型']]
        half = cfg['territory']
        cy0 = fdy + cfg['size']
        verts = [(fdx, cy0 - half), (fdx + half, cy0),
                 (fdx, cy0 + half), (fdx - half, cy0)]
        ax.add_patch(plt.Polygon(verts, facecolor='#EDF2F5',
                                 edgecolor='none', zorder=1.5))

    edge_lines = []
    for gx, gy in territory_cells:
        dx, dy = gx - gy, gx + gy
        if (gx, gy - 1) not in territory_cells:
            edge_lines.append((dx, dy, dx + 1, dy + 1))
        if (gx + 1, gy) not in territory_cells:
            edge_lines.append((dx + 1, dy + 1, dx, dy + 2))
        if (gx, gy + 1) not in territory_cells:
            edge_lines.append((dx, dy + 2, dx - 1, dy + 1))
        if (gx - 1, gy) not in territory_cells:
            edge_lines.append((dx - 1, dy + 1, dx, dy))

    for x1, y1, x2, y2 in edge_lines:
        ax.plot([x1, x2], [y1, y2], color='#5DADE2',
                linewidth=0.15, zorder=1.6)

    for _, b in arc_df.iterrows():
        fdx, fdy = b['dx'], b['dy']
        btype = b['类型']
        cfg = ARC_CFG[btype]
        cx, cy = fdx, fdy + cfg['size']
        if btype == 'flag':
            r = cfg['size'] / np.sqrt(2)
            ax.add_patch(plt.Circle((cx, cy), r, facecolor='#3A6B8A',
                                    edgecolor='none', zorder=5))
        elif btype == 'Headquarters':
            # 用圆代替菱形，避免与相邻玩家的显示菱形重叠
            r = cfg['size'] / np.sqrt(2)
            cx, cy = fdx, fdy + cfg['size']
            ax.add_patch(plt.Circle((cx, cy), r, facecolor='#3A6B8A',
                                    edgecolor='none', zorder=5))

# ========== 玩家渲染（声望深浅 + 火晶金边）==========
def _soften(c, amount=0.45):
    v = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
    v = tuple(int(x + (255 - x) * amount) for x in v)
    return f'#{v[0]:02x}{v[1]:02x}{v[2]:02x}'

def _darken(c, amount=0.3):
    v = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
    v = tuple(int(x * (1 - amount)) for x in v)
    return f'#{v[0]:02x}{v[1]:02x}{v[2]:02x}'

max_p = df['声望'].max()
normal_verts, normal_colors = [], []
fire_verts = []     # 火晶金边菱形

for _, row in df.iterrows():
    gx, gy = row['gx'], row['gy']
    base = row['color']
    prestige = int(row['声望'])
    is_fire = str(row['炉子']).startswith('火')

    # 玩家占据 (gx,gy)~(gx+2,gy+2) 四格，坐标指向底部格子
    v = [to_display(gx, gy), to_display(gx + 2, gy),
         to_display(gx + 2, gy + 2), to_display(gx, gy + 2)]

    if base == NEUTRAL:
        # 非前10联盟：黑色对应全服最高声望，依次减淡
        ratio = min(prestige / max_p, 1.0)
        clr = _soften('#000000', 0.95 - (ratio ** 1.2) * 0.95)
    elif prestige > 10000:
        dark_norm = min((prestige - 10000) / 5000, 1.0)
        clr = _darken(base, dark_norm * 0.3)
    elif prestige <= 3000:
        clr = _soften(base, 0.90)
    else:
        norm = (prestige - 3000) / 7000
        clr = _soften(base, 0.90 - norm * 0.90)

    normal_verts.append(v)
    normal_colors.append(clr)

    if is_fire:
        fire_verts.append(v)

if normal_verts:
    ax.add_collection(PolyCollection(normal_verts, facecolors=normal_colors,
                                     edgecolors='none', zorder=2))
for v in fire_verts:
    ax.add_patch(plt.Polygon(v, fill=False, edgecolor='#FFE033',
                             linewidth=0.15, zorder=2.1))

# ========== 玩家名称（极小字，全图不可见，放大后可见）==========
if SHOW_NAMES:
    for _, row in df.iterrows():
        fs, clr = 5, '#444444'
        ax.text(row['dx'], row['dy'] + 3, row['昵称'],
                fontsize=fs, color=clr,
                ha='center', va='bottom', zorder=3)

# ========== 聚落标注 ==========
for s in settlements:
    dx, dy = s['label']
    ax.text(dx, dy, s['name'], ha='center', va='bottom', fontsize=9,
            color=s['color'], fontweight='bold', zorder=5)

# ========== 熊坑 ==========
if any(s['bear_pit'] for s in settlements):
    bear_img = Image.open('熊坑.png')
    bear_arr = np.array(bear_img)
    mask = (bear_arr[:,:,0] < 30) & (bear_arr[:,:,1] < 30) & (bear_arr[:,:,2] < 30)
    bear_arr = bear_arr.copy()
    bear_arr[mask, 3] = 0
    for s in settlements:
        bp = s['bear_pit']
        if not bp:
            continue
        bx, by = bp
        bp_verts = [
            to_display(bx, by), to_display(bx + 3, by),
            to_display(bx + 3, by + 3), to_display(bx, by + 3),
        ]
        half = 2.25
        cx, cy = bx - by, bx + by + 3
        im = ax.imshow(bear_arr, extent=[cx - half, cx + half, cy - half, cy + half], zorder=6)
        clip_poly = plt.Polygon(bp_verts, transform=ax.transData)
        im.set_clip_path(clip_poly)

# ========== 建筑 ==========
print(f'太阳城: (597, 597)  堡垒: 12')

for name, gx, gy, side in BUILDINGS:
    corners = [
        (gx, gy), (gx + side, gy),
        (gx + side, gy + side), (gx, gy + side),
    ]
    verts = [to_display(x, y) for x, y in corners]
    poly = Polygon(verts, fill=True, facecolor='#e8dcc8',
                   edgecolor='#8b7355', linewidth=1.5, zorder=3)
    ax.add_patch(poly)
    cx, cy = gx + side / 2, gy + side / 2
    dx, dy = to_display(cx, cy)
    fs, fw = (9, 'bold') if name == '太阳城' else (7, 'normal')
    ax.text(dx, dy + 8, name, ha='center', va='bottom', fontsize=fs,
            color='#5c4033', fontweight=fw, zorder=4)

# ========== 图例 ==========
legend_handles = []
for a in alliance_power.head(TOP_N).index:
    c = df.loc[df['联盟'] == a, 'color'].iloc[0]
    cnt = (df['联盟'] == a).sum()
    legend_handles.append(plt.Line2D([0], [0], marker='s', color='w',
                           markerfacecolor=c, markersize=8, label=f'{a} ({cnt})'))
neutral_cnt = (df['color'] == NEUTRAL).sum()
if neutral_cnt > 0:
    legend_handles.append(plt.Line2D([0], [0], marker='s', color='w',
                           markerfacecolor=NEUTRAL, markersize=8,
                           label=f'其他 ({neutral_cnt})'))

ax.legend(handles=legend_handles, loc='upper right', framealpha=0.9,
          facecolor='white', edgecolor='#cccccc', labelcolor='#333',
          fontsize=9, title='联盟', title_fontsize=10)

# ========== 地图边框 ==========
map_border = Polygon([
    to_display(-1, -1), to_display(S, -1),
    to_display(S, S), to_display(-1, S),
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
# 按日期分组，取每日第一份
date_groups = {}
for f in files_all:
    tag = f.replace('3957_', '').replace('.xlsx', '')
    date_key = tag[:4]  # MMDD
    if date_key not in date_groups:  # 取第一份
        date_groups[date_key] = f
daily_dates = sorted(date_groups.keys())
daily_counts = []
for d in daily_dates:
    f = date_groups[d]
    cdf = pd.read_excel(f'{DATA_FOLDER}/{f}')
    daily_counts.append(len(cdf))

ax_tl = fig.add_axes([0.08, 0.75, 0.17, 0.12])  # 左上角
ax_tl.plot(range(len(daily_counts)), daily_counts, 'o-', color='#0F4D92',
           markersize=2.5, linewidth=1.0)
ax_tl.fill_between(range(len(daily_counts)), daily_counts, alpha=0.12, color='#0F4D92')
ax_tl.set_xlim(0, len(daily_counts) - 1)
ax_tl.set_ylim(1500, max(daily_counts) * 1.05)
# 每天标一次
n = len(daily_counts)
tick_labels = [f'{d[:2]}/{d[2:]}' for d in daily_dates]
ax_tl.set_xticks(range(n))
ax_tl.set_xticklabels(tick_labels, fontsize=6, rotation=45)
ax_tl.tick_params(axis='y', labelsize=6)
ax_tl.set_ylabel('人数', fontsize=7)
style_ax(ax_tl)

# --- 左下：炉子等级分布（逐级条形图，Nature风格）---
ax_bl = fig.add_axes([0.08, 0.07, 0.30, 0.11])  # 左下角


def furnace_key(label):
    """炉子排序键：lv.N→N，火N→34+N，火N-M→34+N+M*0.2"""
    if label.startswith('火'):
        s = label.replace('火', '')
        if '-' in s:
            base, sub = s.split('-')
            return 34 + int(base) + int(sub) * 0.2
        return 34 + int(s) if s else 0
    return int(label.replace('lv.', ''))


furnace_counts = {}
for v in df['炉子']:
    s = str(v)
    furnace_counts[s] = furnace_counts.get(s, 0) + 1

sorted_labels = sorted(furnace_counts, key=furnace_key)
filtered = [(l, furnace_counts[l]) for l in sorted_labels if furnace_key(l) >= 7]
all_labels, all_counts = zip(*filtered) if filtered else ([], [])
xp = np.arange(len(all_labels))
ax_bl.bar(xp, all_counts, color='#3775BA', edgecolor='white', linewidth=0.3, width=0.7)
style_ax(ax_bl)
ax_bl.set_xticks(xp)
ax_bl.set_xticklabels(all_labels, fontsize=5.5, rotation=60)
ax_bl.set_xlabel('炉子等级', fontsize=6.5)
ax_bl.set_ylabel('人数', fontsize=6.5)

# --- 右下：声望分布直方图 + 正态拟合（Nature风格）---
ax_br = fig.add_axes([0.66, 0.07, 0.26, 0.10])  # 右下角
pv = df['声望'].values / 10000
mu, std = stats.norm.fit(pv)
bin_edges = np.arange(0, df['声望'].max() + 500, 500) / 10000
counts, bins_out, patches = ax_br.hist(pv, bins=bin_edges, color='#0F4D92',
                                        edgecolor='white', linewidth=0.3, alpha=0.85)
style_ax(ax_br)
ax_br.set_xlim(left=0)
ax_br.set_xlabel('声望（亿）', fontsize=6.5)
ax_br.set_ylabel('人数', fontsize=6.5)
xf = np.linspace(pv.min(), pv.max(), 200)
yf = stats.norm.pdf(xf, mu, std) * len(pv) * (bins_out[1] - bins_out[0])
ax_br.plot(xf, yf, linewidth=1.0, alpha=0.6, color='#B64342')
for bar, c in zip(patches, counts):
    if c > 0:
        ax_br.text(bar.get_x() + bar.get_width()/2, bar.get_height(),
                   f'{int(c)}', ha='center', va='bottom', fontsize=5)

ax.set_xlim(-S - 1, S + 1)
ax.set_ylim(-2, S * 2 + 4)
ax.set_aspect('equal')
ax.axis('off')

fig.savefig(OUTPUT_FILE, dpi=DPI, bbox_inches='tight')
print(f'已保存: {OUTPUT_FILE} (DPI={DPI})')
print(f'玩家总数: {len(df)}')
