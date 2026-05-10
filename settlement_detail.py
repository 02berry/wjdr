"""聚落细节图 — 按聚落切块放大，标注玩家名 + 旗子领地 + 熊坑

输出: Map/{MMDD}/{MMDD}_{联盟}.png (1600DPI, 10×10英寸)

坐标模型: 同 map_view.py（to_display 旋转+缩放）
  玩家坐标 (gx,gy) 指向 2×2 块底部格子
  玩家菱形顶点: (gx,gy)→(gx+2,gy)→(gx+2,gy+2)→(gx,gy+2)

渲染分层:
  zorder 1:   旗子领地填充 + 领地联合边界（格子级，重叠处无线）
  zorder 1.5: 领地合并边界线
  zorder 2:   玩家菱形块（声望深浅 + 灰色网格线）
  zorder 2.1: 火晶金边
  zorder 2.2: 破亿金色方块（现已为深色）
  zorder 3:   玩家名字
  zorder 4:   火晶等级标注
  zorder 5:   旗子圆环 / HQ菱形+文字
  zorder 6:   熊坑图片

着色规则:
  聚落成员 → 联盟色 + 声望深浅
  非成员（路过者）→ 灰色 #ddd0bf
  火晶玩家 → 金边 + 等级数字标注
  声望过亿 → 金色方块（或深色，由配置决定）

聚落检测: 迭代剔除离群点（半径≤20），只输出≥50人聚落
熊坑检测: 扫描 3×3 空格区域，曼哈顿距离选最优

配置区（脚本顶部）:
  DATA_FOLDER, MAP_FOLDER, MAPPING_FILE, S=1200, DPI=1600
  MIN_SETTLEMENT=50, VIEWPORT_PADDING=0.35, VIEWPORT_SCALE=1.0
  NAME_FONTSIZE=5, FILTER_ALLIANCE=None
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
import re, os, warnings
import matplotlib.font_manager as fm
import matplotlib.patheffects as pe
from PIL import Image
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
FILTER_ALLIANCE = None   # 只画指定联盟，None=画全部

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

# ========== 读取建筑表（arc.xlsx，每工作表 = 一个联盟） ==========
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
    before = len(arc_df)
    arc_df = arc_df.drop_duplicates(subset=['gx', 'gy'])
    arc_df['dx'] = arc_df['gx'] - arc_df['gy']
    arc_df['dy'] = arc_df['gx'] + arc_df['gy']
    for t, cfg in ARC_CFG.items():
        cnt = (arc_df['类型'] == t).sum()
        print(f'  {t}: {cnt} 座')
    print(f'  共 {len(arc_df)} 座建筑（去重前{before}，去重后{len(arc_df)}）')
else:
    arc_df = pd.DataFrame()
    print('建筑表不存在，跳过')

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
    if FILTER_ALLIANCE and alliance != FILTER_ALLIANCE:
        continue
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
    # 无 colorbar，不需要留右空间

    # 建筑领地（浅蓝色菱形，zorder=1 在玩家层之下）
    if len(arc_df) > 0:
        for _, b in arc_df.iterrows():
            fdx, fdy = b['dx'], b['dy']
            cfg = ARC_CFG[b['类型']]
            half = cfg['territory']
            cy0 = fdy + cfg['size']
            # 领地包围盒与视口无交则跳过（建筑中心可能在视口外但领地伸入视口）
            if fdx + half < vp[0] or fdx - half > vp[2] or cy0 + half < vp[1] or cy0 - half > vp[3]:
                continue
            territory = [(fdx, cy0 - half), (fdx + half, cy0),
                         (fdx, cy0 + half), (fdx - half, cy0)]
            ax.add_patch(plt.Polygon(territory, facecolor='#EDF2F5',
                                     edgecolor='none', zorder=1))

        # 领地合并边界（游戏格子级联合边界）
        if len(arc_df) > 0:
            # 1. 构建所有领土格子的集合（游戏坐标）
            territory_cells = set()
            for _, b in arc_df.iterrows():
                cfg = ARC_CFG[b['类型']]
                half = cfg['territory']
                cx, cy = b['dx'], b['dy'] + cfg['size']
                fx, fy = b['gx'], b['gy']
                span = half + cfg['size']
                for gx in range(fx - span, fx + span + 1):
                    for gy in range(fy - span, fy + span + 1):
                        dx, dy = gx - gy, gx + gy
                        if abs(dx - cx) + abs(dy + 1 - cy) <= half + 1e-9:
                            territory_cells.add((gx, gy))

            # 2. 筛选在视口附近的格子
            near = set()
            for gx, gy in territory_cells:
                dx, dy = gx - gy, gx + gy
                if dx + 1 >= vp[0] and dx - 1 <= vp[2] and dy + 2 >= vp[1] and dy - 1 <= vp[3]:
                    near.add((gx, gy))

            # 3. 画边界：检查4邻居，不在集合则画共享边
            edge_lines = []
            for gx, gy in near:
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
                        linewidth=0.5, zorder=1.5)

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

    # 建筑标记（旗子圆环 / 总部方块+文字）
    if len(arc_df) > 0:
        c = '#3A6B8A'  # 建筑深蓝
        for _, b in arc_df.iterrows():
            fdx, fdy = b['dx'], b['dy']
            if not (vp[0] <= fdx <= vp[2] and vp[1] <= fdy <= vp[3]):
                continue
            btype = b['类型']
            cfg = ARC_CFG[btype]
            cx, cy = fdx, fdy + cfg['size']

            if btype == 'flag':
                r_outer = cfg['size'] / np.sqrt(2)  # =0.707
                r_inner = r_outer * 0.8
                theta = np.linspace(0, 2 * np.pi, 40)
                px = np.concatenate([cx + r_outer * np.cos(theta),
                                     cx + r_inner * np.cos(theta)[::-1]])
                py = np.concatenate([cy + r_outer * np.sin(theta),
                                     cy + r_inner * np.sin(theta)[::-1]])
                ax.add_patch(plt.Polygon(np.column_stack([px, py]),
                                         facecolor=c, edgecolor='none', zorder=5))
                ax.add_patch(plt.Circle((cx, cy), r_outer, fill=False,
                                        edgecolor=c, linewidth=0.6, zorder=5))
                ax.add_patch(plt.Circle((cx, cy), r_inner, fill=False,
                                        edgecolor=c, linewidth=0.6, zorder=5))

            elif btype == 'Headquarters':
                sz = cfg['size']  # =3
                verts = [(fdx, fdy), (fdx + sz, fdy + sz),
                         (fdx, fdy + sz * 2), (fdx - sz, fdy + sz)]
                ax.add_patch(plt.Polygon(verts, facecolor=c, edgecolor=c,
                                         linewidth=0.6, zorder=5))
                ax.text(cx, cy, cfg['label'], fontsize=8, color='white',
                        ha='center', va='center', fontweight='bold',
                        zorder=6)

    ax.set_xlim(vp[0], vp[2])
    ax.set_ylim(vp[1], vp[3])
    ax.set_aspect('equal')
    ax.axis('off')

    # 标题：左上角
    ax.text(0.005, 0.995, f'{month}.{day} {name} {s["n"]}',
            transform=ax.transAxes, fontsize=6, fontweight='normal',
            ha='left', va='top', color=settle_color.get(name, '#444444'),
            alpha=0.6)

    # 熊坑标记（3×3方块，坐标指向底部格子）
    if s.get('bear_pit'):
        bx, by = s['bear_pit']
        bp_verts = [
            to_display(bx, by),         # 底部
            to_display(bx + 3, by),     # 右
            to_display(bx + 3, by + 3), # 顶部
            to_display(bx, by + 3),     # 左
        ]
        # 熊坑图片（缩小25%，黑色背景透明）
        bear_img = Image.open('熊坑.png')
        bear_arr = np.array(bear_img)
        mask = (bear_arr[:,:,0] < 30) & (bear_arr[:,:,1] < 30) & (bear_arr[:,:,2] < 30)
        bear_arr = bear_arr.copy()
        bear_arr[mask, 3] = 0
        half = 2.25  # 3 * 0.75
        cx, cy = bx - by, bx + by + 3
        im = ax.imshow(bear_arr, extent=[cx - half, cx + half, cy - half, cy + half], zorder=6)
        clip_poly = plt.Polygon(bp_verts, transform=ax.transData)
        im.set_clip_path(clip_poly)

    # 图例已去掉

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
