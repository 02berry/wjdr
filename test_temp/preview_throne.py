"""低DPI预览王城卫星图"""
import sys, os, re, warnings
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.patches import Polygon
import matplotlib.font_manager as fm
warnings.filterwarnings('ignore', message='Glyph')

for fp in ['C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False

DATA_FOLDER = 'data'
MAP_FOLDER = 'Map'
MAPPING_FILE = f'{MAP_FOLDER}/league_mapping.xlsx'
NAME_FONTSIZE = 5

def _soften(c, amount=0.45):
    v = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
    v = tuple(int(x + (255 - x) * amount) for x in v)
    return f'#{v[0]:02x}{v[1]:02x}{v[2]:02x}'

def to_display(gx, gy):
    return gx - gy, gx + gy

def zone_diamond(cx, cy, half):
    corners = [(cx - half, cy - half), (cx + half, cy - half),
               (cx + half, cy + half), (cx - half, cy + half)]
    return [to_display(x, y) for x, y in corners]

files = sorted([f for f in os.listdir(DATA_FOLDER) if f.startswith('3957_') and f.endswith('.xlsx')])
latest_file = os.path.join(DATA_FOLDER, files[-1])
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

lm = pd.read_excel(MAPPING_FILE)
color_map = {}
for _, row in lm.iterrows():
    a, c = row['联盟'], row['颜色']
    if pd.notna(a) and pd.notna(c):
        color_map[str(a).strip()] = c
BG_COLOR = '#f5f2ed'
other_color = '#ddd0bf'
df['联盟'] = df['联盟'].astype(str).str.strip()

MAP_CX, MAP_CY = 599.5, 599.5
ash_half = 14
span_display = 2 * 2 * ash_half * 1.15
half_view = span_display / 2
center_dx, center_dy = to_display(MAP_CX, MAP_CY)
vp_x0, vp_x1 = center_dx - half_view, center_dx + half_view
vp_y0, vp_y1 = center_dy - half_view, center_dy + half_view
in_view = df[(df['dx'] >= vp_x0) & (df['dx'] <= vp_x1) & (df['dy'] >= vp_y0) & (df['dy'] <= vp_y1)]

AUTO_PALETTE = ['#FF8C00', '#9370DB', '#20B2AA', '#D4A017', '#C71585',
                '#4169E1', '#D2691E', '#BA55D3', '#00CED1', '#FF69B4',
                '#8B4513', '#7B68EE', '#00BFFF', '#DB7093', '#6495ED',
                '#CD853F', '#DDA0DD', '#FF7F50', '#4682B4', '#F4A460']
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

fig, ax = plt.subplots(figsize=(10, 10))
fig.patch.set_facecolor(BG_COLOR)
ax.set_facecolor(BG_COLOR)

for name, half, clr in [('沃土', 149.5, '#e8f5e9'), ('废墟', 48.5, '#e0e0e0'), ('灰烬', 14, '#ffcdd2')]:
    ax.add_patch(Polygon(zone_diamond(MAP_CX, MAP_CY, half), fill=True, facecolor=clr, edgecolor='none', zorder=0))

verts = zone_diamond(600, 600, 6.0)
ax.add_patch(Polygon(verts, fill=True, facecolor='#fff3cd', edgecolor='none', zorder=0))

for gx, gy in [(594,594), (604,594), (604,604), (594,604)]:
    cx, cy = to_display(gx + 1, gy + 1)
    ax.add_patch(plt.Circle((cx, cy), radius=1.4, facecolor='none', edgecolor='#c9a800', linewidth=2.0, zorder=2.5))

BUILDINGS = [('1号要塞',597,800,6),('2号要塞',400,597,6),('3号要塞',597,400,6),('4号要塞',800,597,6),
             ('1号堡垒',237,828,6),('2号堡垒',237,606,6),('3号堡垒',237,348,6),('4号堡垒',366,237,6),
             ('5号堡垒',588,237,6),('6号堡垒',846,237,6),('7号堡垒',957,348,6),('8号堡垒',957,606,6),
             ('9号堡垒',957,828,6),('10号堡垒',828,957,6),('11号堡垒',606,957,6),('12号堡垒',348,957,6)]
RUINS_HALF = 29.5
for name, gx, gy, side in BUILDINGS:
    cx, cy = gx + side / 2, gy + side / 2
    dx, dy = to_display(cx, cy)
    if dx < vp_x0 or dx > vp_x1 or dy < vp_y0 or dy > vp_y1:
        continue
    ax.add_patch(Polygon(zone_diamond(cx, cy, RUINS_HALF), fill=True, facecolor='#e0e0e0', edgecolor='none', zorder=1))

verts, colors = [], []
gold_patches = []
fire_labels = []
fire_verts = []

for _, row in in_view.iterrows():
    base = alliance_color.get(str(row['联盟']).strip(), other_color)
    prestige = row['声望']
    is_gold = prestige >= 10000
    if is_gold:
        clr = None
    elif prestige <= 3000:
        clr = _soften(base, 0.90)
    else:
        norm = (prestige - 3000) / 7000
        clr = _soften(base, 0.90 - norm * 0.90)
    furnace = str(row.get('炉子', ''))
    is_fire = furnace.startswith('火')
    if is_fire and not is_gold:
        m = re.search(r'火(\d+)', furnace)
        if m:
            fire_labels.append((row['dx'], row['dy'], m.group(1)))
    v = [to_display(row['gx'], row['gy']), to_display(row['gx']+2, row['gy']),
         to_display(row['gx']+2, row['gy']+2), to_display(row['gx'], row['gy']+2)]
    if is_gold:
        gold_norm = min((prestige - 10000) / 5000, 1.0)
        r = int(0xDA + (0xFF - 0xDA) * gold_norm)
        g = int(0xA5 + (0xD7 - 0xA5) * gold_norm)
        b = int(0x20 + (0x00 - 0x20) * gold_norm)
        gold_patches.append((v, f'#{r:02x}{g:02x}{b:02x}'))
    else:
        verts.append(v)
        colors.append(clr)
    if is_fire and not is_gold:
        fire_verts.append(v)

if verts:
    ax.add_collection(PolyCollection(verts, facecolors=colors, edgecolors='#cccccc', linewidths=0.3, zorder=2))
for v in fire_verts:
    ax.add_patch(plt.Polygon(v, fill=False, edgecolor='#FFE033', linewidth=0.6, zorder=2.1))
for gv, gc in gold_patches:
    ax.add_patch(plt.Polygon(gv, facecolor=gc, edgecolor='none', zorder=2.2))
for dx, dy, lv in fire_labels:
    ax.text(dx, dy + 3.0, lv, fontsize=NAME_FONTSIZE*2.0, color='#FFF44F', ha='center', va='center', fontweight='bold', zorder=4)
for _, row in in_view.iterrows():
    nick = str(row['昵称'])
    if not nick or nick == 'nan':
        continue
    ax.text(row['dx'], row['dy'] + 2, nick, fontsize=NAME_FONTSIZE, color='black', ha='center', va='center', zorder=3)

from PIL import Image as PILImage
from matplotlib.patches import Polygon as MplPolygon
try:
    berry_img = PILImage.open('02.ico').convert('RGBA')
    berry_arr = np.array(berry_img)
    bgx, bgy = 606, 605
    extent = [bgx - bgy - 2, bgx - bgy + 2, bgx + bgy, bgx + bgy + 4]
    berry_im = ax.imshow(berry_arr, extent=extent, aspect='auto', zorder=3.5)
    clip_verts = [to_display(bgx, bgy), to_display(bgx+2, bgy), to_display(bgx+2, bgy+2), to_display(bgx, bgy+2)]
    berry_im.set_clip_path(MplPolygon(clip_verts, transform=ax.transData, closed=True))
except Exception:
    pass

gx, gy, side = 597, 597, 6
corners = [(gx, gy), (gx+side, gy), (gx+side, gy+side), (gx, gy+side)]
verts_b = [to_display(x, y) for x, y in corners]
ax.add_patch(Polygon(verts_b, fill=True, facecolor='#e8dcc8', edgecolor='#8b7355', linewidth=1.5, zorder=4))
ax.set_xlim(vp_x0, vp_x1)
ax.set_ylim(vp_y0, vp_y1)
ax.set_aspect('equal')
ax.axis('off')

ax.text(0.02, 0.97, '3957第三次王战大合照', transform=ax.transAxes, fontsize=16, fontweight='bold', ha='left', va='top', color='#444444')
ax.text(0.02, 0.94, '20260509', transform=ax.transAxes, fontsize=13, fontweight='bold', ha='left', va='top', color='#444444')

legend_handles = []
for alliance in sorted(in_view['联盟'].unique(), key=lambda a: -len(in_view[in_view['联盟']==a])):
    clr = alliance_color.get(alliance, '#888888')
    cnt = len(in_view[in_view['联盟'] == alliance])
    legend_handles.append(plt.Line2D([0], [0], marker='s', color=clr, markerfacecolor=clr, markeredgecolor='none', markersize=8, label=f'{alliance}  {cnt}人'))
ax.legend(handles=legend_handles, loc='lower right', frameon=False, facecolor='white', edgecolor='#cccccc', labelcolor='#333', fontsize=8)

fig.savefig('test_temp/throne_preview.png', dpi=200, bbox_inches='tight')
plt.close(fig)
print('done')
