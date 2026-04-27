import pandas as pd
import re
import warnings
import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import os

warnings.filterwarnings('ignore')

# ==================== 中文字体设置 ====================
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 读取新旧数据 ====================
file_new = 'data/3957_0426.xlsx'
file_old = 'data/3957_0425.xlsx'

df_new = pd.read_excel(file_new, sheet_name='查询结果', dtype=str)
df_old = pd.read_excel(file_old, sheet_name='查询结果', dtype=str)

print(f"今日数据: {len(df_new)} 行")
print(f"昨日数据: {len(df_old)} 行")


# ==================== 数据清洗 ====================
def safe_to_numeric(series, default_value=0):
    result = pd.to_numeric(series, errors='coerce')
    return result.fillna(default_value).astype(float)


df_new['声望'] = safe_to_numeric(df_new['声望'])
df_new['账号'] = safe_to_numeric(df_new['账号'])
df_old['声望'] = safe_to_numeric(df_old['声望'])
df_old['账号'] = safe_to_numeric(df_old['账号'])


def extract_furnace_level(level_str):
    if pd.isna(level_str) or str(level_str).strip() == '':
        return 0
    level_str = str(level_str).strip()
    match = re.search(r'火晶\s*(\d+)', level_str)
    if match:
        return 30 + int(match.group(1))
    match = re.search(r'lv\.?\s*(\d+)', level_str, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 0


df_new['炉子等级'] = df_new['炉子'].apply(extract_furnace_level)
df_old['炉子等级'] = df_old['炉子'].apply(extract_furnace_level)

# ==================== 目标联盟 ====================
TARGET_ALLIANCES = ['FFF', '999', 'KFC']
ALLIANCE_COLORS = {
    'FFF': '#E74C3C',
    '999': '#3498DB',
    'KFC': '#F39C12',
}

# ==================== 构建旧名字字典 ====================
old_name_dict = {}
for _, row in df_old.iterrows():
    account = safe_to_numeric(pd.Series([row['账号']])).iloc[0]
    if account > 0:
        old_name_dict[account] = row['昵称']

# ==================== 声望变化对比 ====================
print("\n=== 声望变化对比 ===")

df_today_target = df_new[df_new['联盟'].str.upper().isin([a.upper() for a in TARGET_ALLIANCES])].copy()

changes = []
for _, row_today in df_today_target.iterrows():
    account = row_today['账号']
    today_prestige = row_today['声望']
    alliance = row_today['联盟'].strip().upper()
    old_name = old_name_dict.get(account, row_today['昵称'])

    old_data = df_old[df_old['账号'] == account]

    if len(old_data) > 0:
        yesterday_prestige = old_data.iloc[0]['声望']
        change = today_prestige - yesterday_prestige
    else:
        yesterday_prestige = None
        change = None

    changes.append({
        '账号': int(account),
        '昵称': old_name,
        '今日昵称': row_today['昵称'],
        '联盟': alliance,
        '炉子': row_today['炉子'],
        '炉子等级': int(row_today['炉子等级']),
        '战前声望': int(yesterday_prestige) if yesterday_prestige is not None else None,
        '战后声望': int(today_prestige),
        '声望变化': int(change) if change is not None else None,
        '坐标': row_today.get('坐标', ''),
        '罩子': row_today.get('罩子', ''),
    })

df_changes = pd.DataFrame(changes)
df_changes_valid = df_changes[df_changes['声望变化'].notna()].copy()
df_changes_valid = df_changes_valid[df_changes_valid['声望变化'] != 0]

df_losses = df_changes_valid[df_changes_valid['声望变化'] < 0].copy()
df_losses['损失量'] = df_losses['声望变化'].abs()
df_losses = df_losses.sort_values('损失量', ascending=False)

# ==================== 绘图数据 ====================
alliance_loss_rankings = {}
alliance_loss_names = {}
for alliance in TARGET_ALLIANCES:
    al_data = df_losses[df_losses['联盟'] == alliance].sort_values('损失量', ascending=False)
    alliance_loss_rankings[alliance] = al_data['损失量'].values
    alliance_loss_names[alliance] = al_data['昵称'].values

max_rank = max(len(v) for v in alliance_loss_rankings.values())
total_people = sum(len(v) for v in alliance_loss_rankings.values())

# ==================== 图1: 战损排名（59人精细版）====================
fig1, ax1 = plt.subplots(figsize=(24, 12), dpi=200)
ax1.set_facecolor('#FAFAFA')
ax1.grid(True, alpha=0.3, linestyle='-', linewidth=0.4, color='#DDDDDD')

# 计算偏移量
y_range = 1381  # 最大损失值
y_offset = y_range * 0.035  # 约48万，够清晰

# 用更粗的线但半透明
for alliance in TARGET_ALLIANCES:
    values = alliance_loss_rankings[alliance]
    names = alliance_loss_names[alliance]
    ranks = np.arange(1, len(values) + 1)
    color = ALLIANCE_COLORS[alliance]

    # 线稍粗但更透明
    ax1.plot(ranks, values, color=color, linewidth=2, alpha=0.25, zorder=1)

    # 点稍大
    ax1.scatter(ranks, values, color=color, s=30, alpha=0.7, zorder=2, edgecolors='white', linewidth=0.5)

    # 逐个标注
    for rank, loss_val, name in zip(ranks, values, names):
        display_name = name if len(name) <= 5 else name[:4] + '…'

        # 🔧 3种位置循环错开
        if rank % 3 == 0:
            # 上方
            offset_y = y_offset
            offset_x = 0
            ha = 'center'
            va = 'bottom'
        elif rank % 3 == 1:
            # 下方
            offset_y = -y_offset
            offset_x = 0
            ha = 'center'
            va = 'top'
        else:
            # 右侧偏移（避免遮挡）
            offset_y = 0
            offset_x = 12
            ha = 'left'
            va = 'center'

        ax1.annotate(display_name,
                     xy=(rank, loss_val),
                     xytext=(offset_x, offset_y),
                     textcoords='offset points',
                     fontsize=9,
                     fontweight='bold',
                     color=color,
                     alpha=0.95,
                     ha=ha,
                     va=va,
                     rotation=0,
                     bbox=dict(boxstyle='round,pad=0.25', facecolor='white',
                               edgecolor=color, alpha=0.9, linewidth=1),
                     zorder=10)

# 图例
from matplotlib.lines import Line2D

legend_elements = [Line2D([0], [0], marker='o', color='w',
                          markerfacecolor=ALLIANCE_COLORS[a],
                          markersize=14, label=f'{a} ({len(alliance_loss_rankings[a])}人)')
                   for a in TARGET_ALLIANCES]
ax1.legend(handles=legend_elements, fontsize=13, loc='upper right',
           framealpha=0.95, edgecolor='#999', title='联盟 (战损人数)')

ax1.set_xlabel('战损排名（从大到小）', fontsize=16, fontweight='bold', labelpad=10)
ax1.set_ylabel('声望损失（万）', fontsize=16, fontweight='bold', labelpad=10)
ax1.set_title('三大联盟 战损排名对比（25号战前昵称）\nFFF 23人 | 999 21人 | KFC 15人',
              fontsize=18, fontweight='bold', pad=20)
ax1.set_xlim(0.5, max_rank + 0.5)
ax1.tick_params(axis='both', labelsize=11)

# 🔧 X轴全部显示1,2,3...（59个不多）
ax1.set_xticks(range(1, max_rank + 1))
ax1.set_xticklabels(range(1, max_rank + 1), fontsize=8, rotation=0)

# Y轴加上"万"单位
ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x)}万'))

plt.tight_layout()
fig1.savefig('战损排名对比.png', dpi=250, bbox_inches='tight', facecolor='white')
print("✅ 图1: 战损排名对比图 已保存（250DPI，59人精细版）")

# ==================== 统计值 ====================
total_losses = []
avg_losses = []
member_counts = []
for alliance in TARGET_ALLIANCES:
    al_data = df_losses[df_losses['联盟'] == alliance]
    total_losses.append(int(al_data['损失量'].sum()) if len(al_data) > 0 else 0)
    avg_losses.append(round(al_data['损失量'].mean(), 1) if len(al_data) > 0 else 0)
    member_counts.append(len(al_data))

max_prestige = max(df_changes_valid['战前声望'].max(), df_changes_valid['战后声望'].max())
min_prestige = min(df_changes_valid['战前声望'].min(), df_changes_valid['战后声望'].min())

# ==================== 图2: 联盟总战损柱状图 ====================
fig2, ax2 = plt.subplots(figsize=(10, 7), dpi=200)

x_pos = np.arange(len(TARGET_ALLIANCES))
bars = ax2.bar(x_pos, total_losses, color=[ALLIANCE_COLORS[a] for a in TARGET_ALLIANCES],
               edgecolor='white', linewidth=2, width=0.55)

for i, (bar, total, count, avg) in enumerate(zip(bars, total_losses, member_counts, avg_losses)):
    height = bar.get_height()
    offset = max(total_losses) * 0.03 if max(total_losses) > 0 else 10
    ax2.text(bar.get_x() + bar.get_width() / 2., height + offset,
             f'总损失: {total}万\n人数: {count} | 人均: {avg}万',
             ha='center', va='bottom', fontsize=13, fontweight='bold',
             color=ALLIANCE_COLORS[TARGET_ALLIANCES[i]])

ax2.set_xticks(x_pos)
ax2.set_xticklabels(TARGET_ALLIANCES, fontsize=15, fontweight='bold')
ax2.set_ylabel('总声望损失（万）', fontsize=14, fontweight='bold')
ax2.set_title('三大联盟 总战损对比', fontsize=17, fontweight='bold', pad=15)
ax2.grid(axis='y', alpha=0.3, linestyle='--')
ax2.tick_params(axis='both', labelsize=12)
if max(total_losses) > 0:
    ax2.set_ylim(0, max(total_losses) * 1.2)

plt.tight_layout()
fig2.savefig('联盟总战损对比.png', dpi=200, bbox_inches='tight', facecolor='white')
print("✅ 图2: 联盟总战损柱状图 已保存（200DPI）")

# ==================== 图3: 声望变化散点图 ====================
fig3, ax3 = plt.subplots(figsize=(12, 10), dpi=200)

ax3.plot([min_prestige, max_prestige], [min_prestige, max_prestige],
         'k--', alpha=0.4, linewidth=2, label='无变化线')
ax3.fill_between([min_prestige, max_prestige], [min_prestige, max_prestige],
                 min_prestige, alpha=0.05, color='red', label='损失区域')
ax3.fill_between([min_prestige, max_prestige], [min_prestige, max_prestige],
                 max_prestige, alpha=0.05, color='green', label='增长区域')

for alliance in TARGET_ALLIANCES:
    color = ALLIANCE_COLORS[alliance]

    al_loss = df_changes_valid[(df_changes_valid['联盟'] == alliance) & (df_changes_valid['声望变化'] < 0)]
    if len(al_loss) > 0:
        ax3.scatter(al_loss['战前声望'], al_loss['战后声望'],
                    c=color, s=al_loss['声望变化'].abs() * 0.8, alpha=0.65,
                    edgecolors='white', linewidth=0.8, label=f'{alliance} 损失')

    al_gain = df_changes_valid[(df_changes_valid['联盟'] == alliance) & (df_changes_valid['声望变化'] > 0)]
    if len(al_gain) > 0:
        ax3.scatter(al_gain['战前声望'], al_gain['战后声望'],
                    c=color, s=al_gain['声望变化'] * 0.8, alpha=0.65,
                    marker='^', edgecolors='white', linewidth=0.8,
                    label=f'{alliance} 增加')

ax3.set_xlabel('战前声望（万）', fontsize=15, fontweight='bold')
ax3.set_ylabel('战后声望（万）', fontsize=15, fontweight='bold')
ax3.set_title('战前 vs 战后 声望变化散点图\n点越大=变化越大 · 线下=损失 · 线上=增长',
              fontsize=16, fontweight='bold', pad=15)
ax3.legend(fontsize=11, loc='upper left', framealpha=0.9, edgecolor='#999')
ax3.grid(True, alpha=0.3, linestyle='--')
ax3.tick_params(axis='both', labelsize=11)
ax3.set_xlim(min_prestige * 0.95, max_prestige * 1.05)
ax3.set_ylim(min_prestige * 0.95, max_prestige * 1.05)

plt.tight_layout()
fig3.savefig('战前战后散点图.png', dpi=200, bbox_inches='tight', facecolor='white')
print("✅ 图3: 战前vs战后散点图 已保存（200DPI）")

# ==================== 图4: 箱线图 ====================
fig4, ax4 = plt.subplots(figsize=(10, 7), dpi=200)

box_data = []
for alliance in TARGET_ALLIANCES:
    al_data = df_losses[df_losses['联盟'] == alliance]['损失量'].values
    box_data.append(al_data if len(al_data) > 0 else [0])

bp = ax4.boxplot(box_data, labels=TARGET_ALLIANCES, patch_artist=True,
                 widths=0.45, showmeans=True,
                 meanprops=dict(marker='D', markerfacecolor='red', markersize=10),
                 medianprops=dict(color='darkred', linewidth=2))

for i, (box, alliance) in enumerate(zip(bp['boxes'], TARGET_ALLIANCES)):
    box.set_facecolor(ALLIANCE_COLORS[alliance])
    box.set_alpha(0.5)

# 散点叠加
for i, alliance in enumerate(TARGET_ALLIANCES):
    al_data = df_losses[df_losses['联盟'] == alliance]['损失量'].values
    if len(al_data) > 0:
        jitter = np.random.normal(0, 0.06, len(al_data))
        ax4.scatter(np.ones(len(al_data)) * (i + 1) + jitter, al_data,
                    alpha=0.4, s=30, color=ALLIANCE_COLORS[alliance], edgecolors='white')

ax4.set_ylabel('声望损失（万）', fontsize=14, fontweight='bold')
ax4.set_title('三大联盟 战损分布箱线图\n红线=均值 · 点=个体', fontsize=16, fontweight='bold', pad=15)
ax4.grid(axis='y', alpha=0.3, linestyle='--')
ax4.tick_params(axis='both', labelsize=12)

plt.tight_layout()
fig4.savefig('战损分布箱线图.png', dpi=200, bbox_inches='tight', facecolor='white')
print("✅ 图4: 战损分布箱线图 已保存（200DPI）")

# ==================== 图5: 综合仪表盘 ====================
fig5 = plt.figure(figsize=(20, 14), dpi=200)
fig5.suptitle('无尽冬日 · 战损综合分析仪表盘', fontsize=22, fontweight='bold', y=0.98)

# 5-1: 左上 - 战损排名曲线
ax5_1 = plt.subplot(2, 3, 1, facecolor='#FAFAFA')
for alliance in TARGET_ALLIANCES:
    values = alliance_loss_rankings[alliance]
    ranks = range(1, len(values) + 1)
    ax5_1.plot(ranks, values, color=ALLIANCE_COLORS[alliance], linewidth=2.5,
               alpha=0.8, label=alliance)
ax5_1.set_title('战损排名曲线', fontsize=14, fontweight='bold')
ax5_1.set_xlabel('排名');
ax5_1.set_ylabel('损失(万)')
ax5_1.legend(fontsize=10);
ax5_1.grid(True, alpha=0.3)
ax5_1.tick_params(labelsize=10)
ax5_1.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))

# 5-2: 中上 - 总战损柱状图
ax5_2 = plt.subplot(2, 3, 2)
bars = ax5_2.bar(TARGET_ALLIANCES, total_losses, color=[ALLIANCE_COLORS[a] for a in TARGET_ALLIANCES],
                 edgecolor='white', linewidth=1.5)
for bar, total in zip(bars, total_losses):
    offset = max(total_losses) * 0.02 if max(total_losses) > 0 else 5
    ax5_2.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + offset,
               f'{total}', ha='center', fontsize=12, fontweight='bold')
ax5_2.set_title('总战损(万)', fontsize=14, fontweight='bold')
ax5_2.grid(axis='y', alpha=0.3)
ax5_2.tick_params(labelsize=11)

# 5-3: 右上 - 饼图
ax5_3 = plt.subplot(2, 3, 3)
if sum(total_losses) > 0:
    colors_pie = [ALLIANCE_COLORS[a] for a in TARGET_ALLIANCES]
    wedges, texts, autotexts = ax5_3.pie(total_losses, labels=TARGET_ALLIANCES,
                                         autopct='%1.1f%%', colors=colors_pie,
                                         startangle=90, explode=[0.04] * 3)
    for autotext in autotexts:
        autotext.set_fontsize(11);
        autotext.set_fontweight('bold')
ax5_3.set_title('战损占比', fontsize=14, fontweight='bold')

# 5-4: 左下 - 散点图
ax5_4 = plt.subplot(2, 3, 4)
for alliance in TARGET_ALLIANCES:
    al_data = df_changes_valid[df_changes_valid['联盟'] == alliance]
    ax5_4.scatter(al_data['战前声望'], al_data['战后声望'],
                  c=ALLIANCE_COLORS[alliance], s=25, alpha=0.5, label=alliance)
if max_prestige > 0:
    ax5_4.plot([min_prestige, max_prestige], [min_prestige, max_prestige], 'k--', alpha=0.4, lw=1.5)
ax5_4.set_xlabel('战前声望');
ax5_4.set_ylabel('战后声望')
ax5_4.set_title('战前 vs 战后', fontsize=14, fontweight='bold')
ax5_4.legend(fontsize=9);
ax5_4.grid(True, alpha=0.3)
ax5_4.tick_params(labelsize=10)

# 5-5: 中下 - 人均战损
ax5_5 = plt.subplot(2, 3, 5)
bars = ax5_5.bar(TARGET_ALLIANCES, avg_losses, color=[ALLIANCE_COLORS[a] for a in TARGET_ALLIANCES],
                 edgecolor='white', linewidth=1.5)
for bar, avg in zip(bars, avg_losses):
    offset = max(avg_losses) * 0.03 if max(avg_losses) > 0 else 5
    ax5_5.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + offset,
               str(avg), ha='center', fontsize=12, fontweight='bold')
ax5_5.set_title('人均损失(万)', fontsize=14, fontweight='bold')
ax5_5.grid(axis='y', alpha=0.3)
ax5_5.tick_params(labelsize=11)

# 5-6: 右下 - 数据汇总表
ax5_6 = plt.subplot(2, 3, 6)
ax5_6.axis('off')
table_data = []
for alliance in TARGET_ALLIANCES:
    al_data = df_losses[df_losses['联盟'] == alliance]
    if len(al_data) > 0:
        table_data.append([alliance, len(al_data), f"{int(al_data['损失量'].sum())}万",
                           f"{int(al_data['损失量'].max())}万", f"{round(al_data['损失量'].mean(), 1)}万"])
    else:
        table_data.append([alliance, 0, '0万', '0万', '0万'])

table = ax5_6.table(cellText=table_data,
                    colLabels=['联盟', '人数', '总损失', '最大损失', '人均损失'],
                    cellLoc='center', loc='center')
table.auto_set_font_size(False)
table.set_fontsize(11)
table.scale(1.1, 2.2)
for i in range(len(TARGET_ALLIANCES)):
    for j in range(5):
        if j == 0:
            table[(i + 1, j)].set_facecolor(ALLIANCE_COLORS[TARGET_ALLIANCES[i]])
            table[(i + 1, j)].set_alpha(0.25)
        table[(i + 1, j)].set_text_props(fontweight='bold')
ax5_6.set_title('数据汇总', fontsize=14, fontweight='bold', y=0.72)

plt.tight_layout()
fig5.savefig('战损综合仪表盘.png', dpi=200, bbox_inches='tight', facecolor='white')
print("✅ 图5: 战损综合仪表盘 已保存（200DPI）")

# ==================== 输出 ====================
print(f"\n{'=' * 60}")
print("📊 全部图表已生成（200 DPI 高分辨率）:")
print("  1. 战损排名对比.png")
print("  2. 联盟总战损对比.png")
print("  3. 战前战后散点图.png")
print("  4. 战损分布箱线图.png")
print("  5. 战损综合仪表盘.png")
print(f"{'=' * 60}")

print(f"\n🔥 战损TOP10:")
if len(df_losses) > 0:
    for i, (_, row) in enumerate(df_losses.head(10).iterrows()):
        name = row['昵称']
        today_name = row['今日昵称']
        name_str = f"{name}" if name == today_name else f"{name}→{today_name}"
        print(f"  {i + 1}. [{row['联盟']}] {name_str} | 损失:{int(row['损失量'])}万")

# ==================== Excel ====================
output_file = '声望战损统计.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    output_cols = ['账号', '昵称', '今日昵称', '联盟', '炉子', '炉子等级',
                   '战前声望', '战后声望', '声望变化', '损失量', '坐标', '罩子']

    if len(df_losses) > 0:
        df_output = df_losses[output_cols].sort_values('损失量', ascending=False)
        df_output.index = range(1, len(df_output) + 1)
        df_output.index.name = '排名'
        df_output.to_excel(writer, sheet_name='战损总榜')

    for alliance in TARGET_ALLIANCES:
        alliance_losses = df_losses[df_losses['联盟'] == alliance]
        if len(alliance_losses) > 0:
            df_al = alliance_losses[output_cols].sort_values('损失量', ascending=False)
            df_al.index = range(1, len(df_al) + 1)
            df_al.index.name = '排名'
            df_al.to_excel(writer, sheet_name=f'{alliance}战损')

    stats_list = []
    for alliance in TARGET_ALLIANCES:
        alliance_data = df_losses[df_losses['联盟'] == alliance]
        stats_list.append({
            '联盟': alliance,
            '战损人数': len(alliance_data),
            '总损失(万)': int(alliance_data['损失量'].sum()) if len(alliance_data) > 0 else 0,
            '人均损失(万)': round(alliance_data['损失量'].mean(), 1) if len(alliance_data) > 0 else 0,
            '最大损失(万)': int(alliance_data['损失量'].max()) if len(alliance_data) > 0 else 0,
        })
    pd.DataFrame(stats_list).to_excel(writer, sheet_name='统计汇总', index=False)

print(f"📁 Excel已保存: {output_file}")