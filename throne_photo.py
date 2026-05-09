"""王城合照图 — 灰烬区域放大，战力映射+火晶/破亿渲染"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.patches import Polygon
import re, os, warnings
import matplotlib.font_manager as fm
import matplotlib.colors as mcolors
import matplotlib.patheffects as pe
from matplotlib.cm import ScalarMappable
warnings.filterwarnings('ignore', message='Glyph')

for fp in ['C:/Windows/Fonts/simsun.ttc', 'C:/Windows/Fonts/times.ttf',
           'C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = ['SimSun', 'Times New Roman']
plt.rcParams['axes.unicode_minus'] = False

# ========== 配置 ==========
DATA_FOLDER = 'data'
MAP_FOLDER = 'Map'
MAPPING_FILE = f'{MAP_FOLDER}/league_mapping.xlsx'
S = 1200
PLAYER_SIZE = 2
DPI = 1600
NAME_FONTSIZE = 5

# 灰烬范围
ASH_HALF = 13.5
ASH_VIEWPORT_MULT = 1.20

# 王域
KINGDOM_HALF = 6.0

# 建筑废墟
RUINS_HALF = 29.5

# 玩家颜色深度映射
GOLD_THRESHOLD = 10_000    # 破亿金色
PRESTIGE_LOW = 3000        # 最低浅色档
PRESTIGE_RANGE = 7000      # 归一化范围（低→高）

# 标题 / 日期 / 右上角文案
TITLE = '3957第三次王战大合照'
TITLE_FONTSIZE = 16
TITLE_X, TITLE_Y = 0.02, 0.97
DATE_FONTSIZE = 13
RIGHT_TEXT = '无尽冬日这么多区，偏偏我们挤进了同一个\n——这大概就是老天安排好的缘分啦'
RIGHT_FONTSIZE = 11
RIGHT_TEXT_X, RIGHT_TEXT_Y = 0.98, 0.97

# 太阳城标注
SUN_CITY_LABEL = '太阳城'
SUN_CITY_FONTSIZE = 16

# 玩家头像
AVATAR_FILE = '纯白头像.png'
AVATAR_COORDS = (606, 605)  # (gx, gy)

# 颜色
BG_COLOR = '#f5f2ed'
OTHER_COLOR = '#ddd0bf'

def display_width(s):
    w = 0
    for c in s:
        w += 2 if ord(c) > 127 else 1
    return w

def split_name(s, max_w=10):
    if display_width(s) <= max_w:
        return s
    half = display_width(s) // 2
    cum = 0
    for i, c in enumerate(s):
        cum += 2 if ord(c) > 127 else 1
        if cum >= half:
            return s[:i] + '\n' + s[i:]
    return s

def _soften(c, amount=0.45):
    v = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
    v = tuple(int(x + (255 - x) * amount) for x in v)
    return f'#{v[0]:02x}{v[1]:02x}{v[2]:02x}'

def to_display(gx, gy):
    return gx - gy, gx + gy

def zone_diamond(cx, cy, half):
    corners = [
        (cx - half, cy - half), (cx + half, cy - half),
        (cx + half, cy + half), (cx - half, cy + half),
    ]
    return [to_display(x, y) for x, y in corners]

# ========== 读取最新数据 ==========
files = sorted([f for f in os.listdir(DATA_FOLDER)
                if f.startswith('3957_') and f.endswith('.xlsx')])
latest_file = os.path.join(DATA_FOLDER, files[-1])
print(f'读取: {latest_file}')
date_tag = files[-1].replace('3957_', '').replace('.xlsx', '')
date_str = re.sub(r'[a-zA-Z]', '', date_tag)

df = pd.read_excel(latest_file)

coords = []
for c in df['坐标']:
    m = re.search(r'\((\d+),(\d+)\)', str(c))
    coords.append((int(m.group(1)), int(m.group(2))) if m else (None, None))
df['gx'], df['gy'] = zip(*coords)
df = df.dropna(subset=['gx', 'gy'])
df[['gx', 'gy']] = df[['gx', 'gy']].astype(int)
df['dx'] = df['gx'] - df['gy']
df['dy'] = df['gx'] + df['gy']

# ========== 联盟颜色 ==========
lm = pd.read_excel(MAPPING_FILE)
color_map = {}
for _, row in lm.iterrows():
    a, c = row['联盟'], row['颜色']
    if pd.notna(a) and pd.notna(c):
        color_map[str(a).strip()] = c

df['联盟'] = df['联盟'].astype(str).str.strip()

# ========== 地形 ==========
MAP_CX, MAP_CY = 599.5, 599.5
THRONE_ZONES = [
    ('沃土', 149.5, '#e8f5e9'),
    ('废墟',  48.5, '#e0e0e0'),
    ('灰烬',  ASH_HALF, '#ffcdd2'),
]

KINGDOM_CX, KINGDOM_CY = 600, 600

# ========== 建筑 ==========
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

# ========== 视口（ASH_VIEWPORT_MULT×灰烬）==========
ash_half = ASH_HALF
span_display = 2 * 2 * ash_half * ASH_VIEWPORT_MULT
half_view = span_display / 2
center_dx, center_dy = to_display(MAP_CX, MAP_CY)
vp_x0 = center_dx - half_view
vp_x1 = center_dx + half_view
vp_y0 = center_dy - half_view
vp_y1 = center_dy + half_view
print(f'视口: dx [{vp_x0:.0f}, {vp_x1:.0f}]  dy [{vp_y0:.0f}, {vp_y1:.0f}]')

in_view = df[(df['dx'] >= vp_x0) & (df['dx'] <= vp_x1) &
             (df['dy'] >= vp_y0) & (df['dy'] <= vp_y1)]
print(f'视野内玩家: {len(in_view)}')

# ========== 为视野内所有联盟分配颜色 ==========
AUTO_PALETTE = [
    '#FF8C00', '#9370DB', '#20B2AA', '#D4A017', '#C71585',
    '#4169E1', '#D2691E', '#BA55D3', '#00CED1', '#FF69B4',
    '#8B4513', '#7B68EE', '#00BFFF', '#DB7093', '#6495ED',
    '#CD853F', '#DDA0DD', '#FF7F50', '#4682B4', '#F4A460',
]
used_colors = set(color_map.values())
auto_idx = 0
alliance_color = {}
for alliance in in_view['联盟'].unique():
    if alliance in color_map:
        alliance_color[alliance] = color_map[alliance]
    else:
        while auto_idx < len(AUTO_PALETTE) and AUTO_PALETTE[auto_idx] in used_colors:
            auto_idx += 1
        clr = AUTO_PALETTE[auto_idx % len(AUTO_PALETTE)]
        auto_idx += 1
        used_colors.add(clr)
        alliance_color[alliance] = clr

# ========== 绘图 ==========
fig, ax = plt.subplots(figsize=(10, 10))
fig.patch.set_facecolor(BG_COLOR)
ax.set_facecolor(BG_COLOR)
# 地形
for name, half, clr in THRONE_ZONES:
    verts = zone_diamond(MAP_CX, MAP_CY, half)
    ax.add_patch(Polygon(verts, fill=True, facecolor=clr,
                         edgecolor='none', zorder=0))

# 王域（太阳城专属地形，中心 (600,600)）
verts = zone_diamond(KINGDOM_CX, KINGDOM_CY, KINGDOM_HALF)
ax.add_patch(Polygon(verts, fill=True, facecolor='#fff3cd',
                     edgecolor='none', zorder=0))

# 炮台（王域四角，1×1暗色圆环）
# 炮台（王域四角，2×2方块暗色圆环）
for gx, gy in [(594,594), (604,594), (604,604), (594,604)]:
    cx, cy = to_display(gx + 1, gy + 1)
    ax.add_patch(plt.Circle((cx, cy), radius=1.4, facecolor='none',
                            edgecolor='#c9a800', linewidth=2.0, zorder=2.5))

# 建筑废墟（在视野内的才画）
for name, gx, gy, side in BUILDINGS:
    if '要塞' not in name and '堡垒' not in name:
        continue
    cx, cy = gx + side / 2, gy + side / 2
    dx, dy = to_display(cx, cy)
    if dx < vp_x0 or dx > vp_x1 or dy < vp_y0 or dy > vp_y1:
        continue
    verts = zone_diamond(cx, cy, RUINS_HALF)
    ax.add_patch(Polygon(verts, fill=True, facecolor='#e0e0e0',
                         edgecolor='none', zorder=1))

# ========== 玩家渲染（同聚落逻辑）==========
verts, colors = [], []
gold_patches = []   # 过亿：(菱形顶点, 金色)
fire_labels = []    # 火晶：(dx, dy, 等级数字)
fire_verts = []     # 火晶金边

for _, row in in_view.iterrows():
    base = alliance_color.get(str(row["联盟"]).strip(), OTHER_COLOR)
    prestige = row['声望']
    is_gold = prestige >= GOLD_THRESHOLD

    if is_gold:
        clr = None
    elif prestige <= PRESTIGE_LOW:
        clr = _soften(base, 0.90)
    else:
        norm = (prestige - PRESTIGE_LOW) / PRESTIGE_RANGE
        clr = _soften(base, 0.90 - norm * 0.90)

    furnace = str(row.get('炉子', ''))
    is_fire = furnace.startswith('火')

    if is_fire and not is_gold:
        m = re.search(r'火(\d+)', furnace)
        if m:
            fire_labels.append((row['dx'], row['dy'], m.group(1)))

    v = [
        to_display(row['gx'], row['gy']),
        to_display(row['gx'] + 2, row['gy']),
        to_display(row['gx'] + 2, row['gy'] + 2),
        to_display(row['gx'], row['gy'] + 2),
    ]

    if is_gold:
        gold_norm = min((prestige - GOLD_THRESHOLD) / 5000, 1.0)
        r = int(0xDA + (0xFF - 0xDA) * gold_norm)
        g = int(0xA5 + (0xD7 - 0xA5) * gold_norm)
        b = int(0x20 + (0x00 - 0x20) * gold_norm)
        gold_patches.append((v, f'#{r:02x}{g:02x}{b:02x}'))
    else:
        verts.append(v)
        colors.append(clr)

    if is_fire and not is_gold:
        fire_verts.append(v)

# 主方块层 + 灰网格线
if verts:
    ax.add_collection(PolyCollection(verts, facecolors=colors,
                                     edgecolors='#cccccc',
                                     linewidths=0.3, zorder=2))

# 火晶金边
for v in fire_verts:
    ax.add_patch(plt.Polygon(v, fill=False, edgecolor='#FFE033',
                             linewidth=0.6, zorder=2.1))

# 过亿金色方块（无边）
for gv, gc in gold_patches:
    ax.add_patch(plt.Polygon(gv, facecolor=gc, edgecolor='none', zorder=2.2))

# 火晶等级数字
for dx, dy, lv in fire_labels:
    ax.text(dx, dy + 3.0, lv, fontsize=NAME_FONTSIZE * 2.0,
            color='#FFF44F', ha='center', va='center',
            fontweight='bold', zorder=4)

# 玩家名字
for _, row in in_view.iterrows():
    nick = str(row['昵称'])
    if not nick or nick == 'nan':
        continue
    ax.text(row['dx'], row['dy'] + 2, split_name(nick),
            fontsize=NAME_FONTSIZE, color='black',
            ha='center', va='center', zorder=3)

# 02berry 头像标记
try:
    from PIL import Image as PILImage
    from matplotlib.patches import Polygon as MplPolygon
    import numpy as np
    berry_img = PILImage.open(AVATAR_FILE).convert('RGBA')
    berry_arr = np.array(berry_img)
    bgx, bgy = AVATAR_COORDS
    extent = [bgx - bgy - 2, bgx - bgy + 2, bgx + bgy, bgx + bgy + 4]
    berry_im = ax.imshow(berry_arr, extent=extent, aspect='auto', zorder=3.5)
    # 用菱形clip_path裁切，避免遮到相邻玩家
    clip_verts = [to_display(bgx, bgy), to_display(bgx+2, bgy),
                  to_display(bgx+2, bgy+2), to_display(bgx, bgy+2)]
    berry_im.set_clip_path(MplPolygon(clip_verts, transform=ax.transData, closed=True))
    print('02berry 头像已加载（纯白头像.png）')
except Exception as e:
    print(f'02berry 头像加载失败: {e}')

# 太阳城建筑
gx, gy, side = 597, 597, 6
corners = [
    (gx, gy), (gx + side, gy),
    (gx + side, gy + side), (gx, gy + side),
]
verts_b = [to_display(x, y) for x, y in corners]
ax.add_patch(Polygon(verts_b, fill=True, facecolor='#e8dcc8',
                     edgecolor='#8b7355', linewidth=1.5, zorder=4))
ax.set_xlim(vp_x0, vp_x1)
ax.set_ylim(vp_y0, vp_y1)
ax.set_aspect('equal')
ax.axis('off')

# 标题（左对齐保持原位）
title_obj = ax.text(TITLE_X, TITLE_Y, TITLE,
        transform=ax.transAxes, fontsize=TITLE_FONTSIZE, fontweight='bold',
        ha='left', va='top', color='#444444')

# 测量大合照宽度，日期与大合照右对齐
fig.canvas.draw()
renderer = fig.canvas.get_renderer()
bbox = title_obj.get_window_extent(renderer)
inv = ax.transAxes.inverted()
title_right = inv.transform((bbox.x1, 0))[0]
ax.text(title_right, 0.932, '20260509',
        transform=ax.transAxes, fontsize=DATE_FONTSIZE, fontweight='bold',
        ha='right', va='top', color='#444444')

# 右上角文案
ax.text(RIGHT_TEXT_X, RIGHT_TEXT_Y, RIGHT_TEXT,
        transform=ax.transAxes, fontsize=RIGHT_FONTSIZE, ha='right', va='top',
        color='#444444', linespacing=1.5)

# 太阳城文字（写在建筑方块内）
sc_dx, sc_dy = to_display(600, 600)
ax.text(sc_dx, sc_dy, SUN_CITY_LABEL, fontsize=SUN_CITY_FONTSIZE, fontweight='bold',
        color='#5c4033', ha='center', va='center', zorder=5)

# 图例（整体右移，色块移至数字后方）
sorted_ali = sorted(in_view['联盟'].unique(), key=lambda a: len(in_view[in_view['联盟']==a]))
top_ali = sorted_ali[-1]
LX_C = 0.73  # 色条起点
LX_N = 0.78  # 文字起点
LY0 = 0.02
LY_STEP = 0.025
# 用渲染器量第一行联盟名宽度，确定人数列位置
t = ax.text(0, 0, top_ali, fontsize=7, transform=ax.transAxes)
bb = t.get_window_extent(renderer)
t.remove()
del t
name_r = inv.transform((bb.x1, 0))[0]
cx = LX_N + name_r + 0.025  # 联盟名右端 + 4空格 = 人数列
# 整体右移：色块移到数字后面 + 再移一次2/5
shift = cx - LX_C
total = shift * 1.4  # 1 + 2/5
LX_C += total
LX_N += total
cx += total
for i, ali in enumerate(sorted_ali):
    clr = alliance_color.get(ali, '#888888')
    cnt = len(in_view[in_view['联盟'] == ali])
    y = LY0 + i * LY_STEP
    ax.plot([LX_C, LX_C + 0.025], [y, y], color=clr, linewidth=3,
            transform=ax.transAxes, clip_on=False)
    ax.text(LX_N, y, ali, transform=ax.transAxes,
            fontsize=7, va='center', color='#333')
    ax.text(cx, y, f'{cnt}', transform=ax.transAxes,
            fontsize=7, va='center', ha='center', color='#333')

# 输出
out_dir = f'{MAP_FOLDER}/{date_str}'
os.makedirs(out_dir, exist_ok=True)
out = f'{out_dir}/{date_str}_王城合照.png'
fig.savefig(out, dpi=DPI, bbox_inches='tight')
print(f'已保存: {out}')
plt.close(fig)
