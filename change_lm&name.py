import pandas as pd
import re
import os
import glob
import warnings
import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime

warnings.filterwarnings('ignore')
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 🔧 配置区（按需修改） ====================

# 数据文件夹路径
DATA_FOLDER = 'data'

# 文件名匹配模式
FILE_PATTERN = '3957_*.xlsx'

# 图表筛选：声望大于多少才画图
MIN_PRESTIGE_FOR_PLOT = 3000

# 图表输出数量限制
TOP_N = 40

# ==================== 🛡️ 豁免账号配置 ====================
EXEMPT_ACCOUNTS = [
    904754949,  # 盟主账号
]

# ==================== 🎨 联盟颜色配置 ====================
RED_ALLIANCES = ['FFF', '666', 'OoO', 'SSR']
RED_COLOR = '#E74C3C'

GREEN_ALLIANCES = ['999', 'ann', 'xxx']
GREEN_COLOR = '#27AE60'

BLUE_ALLIANCES = ['KFC']
BLUE_COLOR = '#3498DB'

DEFAULT_COLOR = '#000000'

ALLIANCE_COLORS = {}
for a in RED_ALLIANCES:
    ALLIANCE_COLORS[a.upper()] = RED_COLOR
for a in GREEN_ALLIANCES:
    ALLIANCE_COLORS[a.upper()] = GREEN_COLOR
for a in BLUE_ALLIANCES:
    ALLIANCE_COLORS[a.upper()] = BLUE_COLOR

LEGEND_LABELS = {
    RED_COLOR: '/'.join(RED_ALLIANCES),
    GREEN_COLOR: '/'.join(GREEN_ALLIANCES),
    BLUE_COLOR: '/'.join(BLUE_ALLIANCES),
}


def get_alliance_color(alliance):
    if pd.isna(alliance) or alliance == '' or alliance == '-':
        return DEFAULT_COLOR
    return ALLIANCE_COLORS.get(alliance.strip().upper(), DEFAULT_COLOR)


# ==================== 查找文件 ====================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_PATH = os.path.join(CURRENT_DIR, DATA_FOLDER)

all_files = glob.glob(os.path.join(FOLDER_PATH, FILE_PATTERN))
all_files = sorted(all_files)
print(f"找到 {len(all_files)} 个文件:")

file_times = []
for f in all_files:
    basename = os.path.basename(f)
    match = re.search(r'3957_(\d{4})', basename)
    if match:
        time_str = match.group(1)
        file_times.append((f, time_str))
        print(f"  - {basename} → {time_str}")

if len(file_times) == 0:
    print("❌ 无文件")
    exit()


def time_to_date(time_str):
    month = int(time_str[:2])
    day = int(time_str[2:])
    return datetime(2026, month, day)


time_dates = [time_to_date(t) for _, t in file_times]
time_order = [t for _, t in file_times]

date_diffs = []
for i in range(len(time_dates)):
    if i == 0:
        date_diffs.append(0)
    else:
        diff = (time_dates[i] - time_dates[i - 1]).days
        if diff <= 1:
            date_diffs.append(1.0)
        elif diff == 2:
            date_diffs.append(1.2)
        elif diff == 3:
            date_diffs.append(1.4)
        else:
            date_diffs.append(1.5)

cumulative_days = np.cumsum(date_diffs)
total_days_span = cumulative_days[-1] - cumulative_days[0]


# ==================== 读取数据 ====================
def safe_to_numeric(series, default_value=0):
    result = pd.to_numeric(series, errors='coerce')
    return result.fillna(default_value).astype(float)


account_timeline = {}

for file_path, time_str in file_times:
    df = pd.read_excel(file_path, sheet_name=0, dtype=str)
    df['账号'] = safe_to_numeric(df['账号'])
    df['声望'] = safe_to_numeric(df['声望'])

    for _, row in df.iterrows():
        account = int(row['账号'])
        if account == 0:
            continue
        if account in EXEMPT_ACCOUNTS:
            continue

        alliance = str(row['联盟']).strip() if pd.notna(row['联盟']) else '-'
        name = str(row['昵称']).strip() if pd.notna(row['昵称']) else ''
        prestige = int(row['声望'])

        if account not in account_timeline:
            account_timeline[account] = {}

        account_timeline[account][time_str] = {
            '昵称': name,
            '联盟': alliance,
            '声望': prestige,
        }

print(f"追踪 {len(account_timeline)} 个账号（已排除豁免）")

# ==================== 分析变化 ====================
change_records = []

for account, timeline in account_timeline.items():
    sorted_times = sorted(timeline.keys())
    if len(sorted_times) < 2:
        continue

    name_changes = []
    alliance_changes = []

    prev_name = None
    prev_alliance = None

    for t in sorted_times:
        current_name = timeline[t]['昵称']
        current_alliance = timeline[t]['联盟']

        if prev_name is not None and current_name != prev_name:
            name_changes.append((t, prev_name, current_name))

        if prev_alliance is not None and current_alliance != prev_alliance:
            if prev_alliance != '-' and current_alliance != '-':
                alliance_changes.append((t, prev_alliance, current_alliance))
            elif prev_alliance != '-' and current_alliance == '-':
                pass
            elif prev_alliance == '-' and current_alliance != '-':
                last_real_alliance = None
                for pt in reversed(sorted_times[:sorted_times.index(t)]):
                    if timeline[pt]['联盟'] != '-':
                        last_real_alliance = timeline[pt]['联盟']
                        break
                if last_real_alliance:
                    alliance_changes.append((t, last_real_alliance, current_alliance))

        prev_name = current_name
        prev_alliance = current_alliance

    total_changes = len(name_changes) + len(alliance_changes)

    if total_changes > 0:
        latest = timeline[sorted_times[-1]]

        name_change_summary = ' → '.join([f"{old}→{new}" for _, old, new in name_changes]) if name_changes else '无'
        alliance_summary_parts = [f"{old}→{new}" for _, old, new in alliance_changes]
        alliance_change_summary = ' | '.join(alliance_summary_parts) if alliance_summary_parts else '无'

        change_dates = set()
        for t_data in name_changes:
            change_dates.add(t_data[0])
        for t_data in alliance_changes:
            change_dates.add(t_data[0])
        if len(sorted_times) > 0:
            change_dates.add(sorted_times[0])
            change_dates.add(sorted_times[-1])

        change_records.append({
            '账号': account,
            '昵称': latest['昵称'],
            '联盟': latest['联盟'],
            '声望': latest['声望'],
            '昵称变化次数': len(name_changes),
            '联盟变化次数': len(alliance_changes),
            '总变化次数': total_changes,
            '昵称变化': name_change_summary,
            '联盟变化': alliance_change_summary,
            '有联盟变化': len(alliance_changes) > 0,
            '_timeline': timeline,
            '_sorted_times': sorted_times,
            '_name_changes': name_changes,
            '_alliance_changes': alliance_changes,
            '_change_dates': change_dates,
        })

df_changes = pd.DataFrame(change_records)

if len(df_changes) > 0:
    df_alliance_changed = df_changes[df_changes['有联盟变化']].sort_values(
        ['联盟变化次数', '声望'], ascending=[False, False])
    df_name_only = df_changes[~df_changes['有联盟变化']].sort_values(
        ['昵称变化次数', '声望'], ascending=[False, False])
    df_changes = pd.concat([df_alliance_changed, df_name_only], ignore_index=True)
    df_changes.index = range(1, len(df_changes) + 1)
    df_changes.index.name = '序号'
else:
    print("❌ 无变化")
    exit()

# ==================== 保存Excel ====================
output_excel = '昵称联盟变化分析.xlsx'
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    output_cols = ['账号', '昵称', '联盟', '声望', '昵称变化次数', '联盟变化次数',
                   '总变化次数', '昵称变化', '联盟变化']
    df_changes[output_cols].to_excel(writer, sheet_name='变化总览', index=True)
    if len(df_alliance_changed) > 0:
        df_alliance_changed[output_cols].to_excel(writer, sheet_name='联盟变化', index=True)
print(f"📁 {output_excel}")


# ==================== 通用绘图函数 ====================
def draw_timeline(df_plot_full, output_name, title_suffix, sort_by='alliance'):
    """绘制时序图"""
    df_plot = df_plot_full.head(TOP_N).copy()
    if len(df_plot) == 0:
        print(f"⚠️ 无符合条件的玩家，跳过 {output_name}")
        return

    n_players = len(df_plot)

    fig_height = max(8, n_players * 0.45)
    fig_width = max(14, min(40, total_days_span * 0.5))

    if n_players > 20 or total_days_span > 20:
        label_fontsize = 6
        point_size = 18
        line_width = 1.5
        stat_fontsize = 5.5
    elif n_players > 10 or total_days_span > 10:
        label_fontsize = 7
        point_size = 22
        line_width = 2
        stat_fontsize = 6.5
    else:
        label_fontsize = 8
        point_size = 28
        line_width = 2.5
        stat_fontsize = 7

    bar_max_length = total_days_span * 0.06

    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=150)

    all_prestige = [row['声望'] for _, row in df_plot.iterrows()]
    max_prestige = max(all_prestige) if all_prestige else 1
    min_prestige = min(all_prestige) if all_prestige else 0
    prestige_range = max_prestige - min_prestige if max_prestige > min_prestige else 1

    for i, (_, row) in enumerate(df_plot.iterrows()):
        y_pos = i + 1
        timeline = row['_timeline']
        sorted_times = row['_sorted_times']
        change_dates = row['_change_dates']

        for j in range(len(sorted_times) - 1):
            t_curr = sorted_times[j]
            t_next = sorted_times[j + 1]
            x_curr = cumulative_days[time_order.index(t_curr)]
            x_next = cumulative_days[time_order.index(t_next)]
            alliance_curr = timeline[t_curr]['联盟']

            if alliance_curr == '-':
                color = '#CCCCCC'
                alpha = 0.3
                linestyle = ':'
            else:
                color = get_alliance_color(alliance_curr)
                alpha = 0.6
                linestyle = '-'

            ax.plot([x_curr, x_next], [y_pos, y_pos], color=color, linewidth=line_width,
                    alpha=alpha, linestyle=linestyle, zorder=1)

        for t in sorted_times:
            if t not in change_dates:
                continue
            x = cumulative_days[time_order.index(t)]
            alliance = timeline[t]['联盟']

            # 🔧 修复：无联盟空心圆，用 edgecolors 和 facecolors
            if alliance == '-':
                ax.scatter(x, y_pos, s=point_size, facecolors='none', edgecolors='#999999',
                           linewidths=1.2, alpha=0.6, zorder=3)
            else:
                ax.scatter(x, y_pos, s=point_size, c=get_alliance_color(alliance), alpha=0.85,
                           edgecolors='white', linewidths=0.5, zorder=3)

        prev_name = timeline[sorted_times[0]]['昵称']
        prev_alliance = timeline[sorted_times[0]]['联盟']

        for j, t in enumerate(sorted_times):
            if j == 0:
                continue
            x = cumulative_days[time_order.index(t)]
            current_name = timeline[t]['昵称']
            current_alliance = timeline[t]['联盟']
            annotations_here = []

            if current_alliance != prev_alliance and prev_alliance != '-' and current_alliance != '-':
                annotations_here.append(('alliance', f"{prev_alliance}→{current_alliance}"))
            if current_name != prev_name:
                annotations_here.append(('name', f"{prev_name}→{current_name}"))

            n_ann = len(annotations_here)
            for k, (ann_type, ann_text) in enumerate(annotations_here):
                if n_ann == 2:
                    offset_y = 14 if ann_type == 'alliance' else 2
                else:
                    offset_y = 10

                ax.annotate(ann_text,
                            xy=(x, y_pos), xytext=(0, offset_y),
                            textcoords='offset points', fontsize=label_fontsize, fontweight='bold',
                            ha='center', va='bottom',
                            color=get_alliance_color(current_alliance) if ann_type == 'alliance' else '#555',
                            zorder=10)

            prev_name = current_name
            prev_alliance = current_alliance

        first_t = sorted_times[0]
        x_first = cumulative_days[time_order.index(first_t)]
        first_alliance = timeline[first_t]['联盟']
        first_name = timeline[first_t]['昵称']
        first_color = get_alliance_color(first_alliance) if first_alliance != '-' else '#999999'

        ax.annotate(f"({first_alliance}){first_name}",
                    xy=(x_first, y_pos),
                    xytext=(-10, 0),
                    textcoords='offset points',
                    fontsize=label_fontsize - 1, color=first_color,
                    ha='right', va='center', fontweight='bold', zorder=8)

        x_last = cumulative_days[time_order.index(sorted_times[-1])]
        prestige = row['声望']
        stats_text = f"{row['联盟变化次数']}+{row['昵称变化次数']}" if sort_by == 'alliance' else f"{row['昵称变化次数']}+{row['联盟变化次数']}"

        if prestige_range > 0:
            bar_length = bar_max_length * (prestige - min_prestige) / prestige_range + bar_max_length * 0.1
        else:
            bar_length = bar_max_length * 0.3

        text_x = x_last + total_days_span * 0.04
        bar_start_x = text_x + total_days_span * 0.07

        ax.text(text_x, y_pos, stats_text, fontsize=stat_fontsize, va='center', ha='left',
                fontweight='bold', color='#555')

        bar_color = get_alliance_color(timeline[sorted_times[-1]]['联盟'])
        if timeline[sorted_times[-1]]['联盟'] == '-':
            bar_color = '#999999'

        ax.barh(y_pos, bar_length, left=bar_start_x, height=0.35,
                color=bar_color, alpha=0.7, edgecolor='white', linewidth=0.3, zorder=3)
        ax.text(bar_start_x + bar_length + total_days_span * 0.01, y_pos, str(prestige),
                fontsize=stat_fontsize - 1, va='center', ha='left', color='#555', fontweight='bold')

    # X轴和X轴标题保留
    ax.set_xticks(cumulative_days)
    if total_days_span > 30:
        xtick_fontsize = 6
        rotation = 45
    elif total_days_span > 15:
        xtick_fontsize = 8
        rotation = 30
    else:
        xtick_fontsize = 10
        rotation = 0
    ax.set_xticklabels(time_order, fontsize=xtick_fontsize, rotation=rotation)
    ax.set_xlabel('日期', fontsize=13, fontweight='bold')

    # 🔧 去掉上、左、右边框和Y轴标题
    ax.spines['top'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(left=False, labelleft=False)
    ax.set_ylabel('')

    left_margin = total_days_span * 0.06
    right_margin = total_days_span * 0.18
    ax.set_xlim(cumulative_days[0] - left_margin, cumulative_days[-1] + right_margin)
    ax.set_ylim(n_players + 1.5, 0.5)

    # 去掉标题
    # (不设置任何标题)

    # 🔧 图例放右上角
    from matplotlib.lines import Line2D
    legend_elements = []
    if RED_ALLIANCES:
        legend_elements.append(Line2D([0], [0], marker='o', color='w', markerfacecolor=RED_COLOR,
                                      markersize=8, label=LEGEND_LABELS[RED_COLOR]))
    if GREEN_ALLIANCES:
        legend_elements.append(Line2D([0], [0], marker='o', color='w', markerfacecolor=GREEN_COLOR,
                                      markersize=8, label=LEGEND_LABELS[GREEN_COLOR]))
    if BLUE_ALLIANCES:
        legend_elements.append(Line2D([0], [0], marker='o', color='w', markerfacecolor=BLUE_COLOR,
                                      markersize=8, label=LEGEND_LABELS[BLUE_COLOR]))
    legend_elements.append(Line2D([0], [0], marker='o', color='w', markerfacecolor=DEFAULT_COLOR,
                                  markersize=8, label='其他'))
    legend_elements.append(Line2D([0], [0], marker='o', color='w',
                                  markerfacecolor='none', markeredgecolor='#999999',
                                  markersize=8, markeredgewidth=1.5, label='无联盟'))

    # 🔧 图例字：用"战力"替代"声望"，去掉横线符号
    if sort_by == 'alliance':
        legend_text = '联盟变化+昵称变化 战力'
    else:
        legend_text = '昵称变化+联盟变化 战力'
    legend_elements.append(Line2D([0], [0], color='white', linewidth=0, label=legend_text))

    ax.legend(handles=legend_elements, fontsize=7, loc='upper right',
              framealpha=0.9, edgecolor='#ccc', ncol=1)

    plt.tight_layout()
    fig.savefig(output_name, dpi=150, bbox_inches='tight', facecolor='white')
    print(f"✅ {output_name} 已保存 ({n_players}/{len(df_plot_full)}个玩家)")


# ==================== 准备绘图数据 ====================
df_plot1 = df_changes[(df_changes['有联盟变化']) & (df_changes['声望'] > MIN_PRESTIGE_FOR_PLOT)].copy()
df_plot1 = df_plot1.sort_values(['联盟变化次数', '声望'], ascending=[False, False]).reset_index(drop=True)

df_plot2 = df_changes[(df_changes['昵称变化次数'] > 0) & (df_changes['声望'] > MIN_PRESTIGE_FOR_PLOT)].copy()
df_plot2 = df_plot2.sort_values(['昵称变化次数', '联盟变化次数', '声望'], ascending=[False, False, False]).reset_index(drop=True)


# ==================== 生成两张图 ====================
if len(df_plot1) > 0:
    draw_timeline(df_plot1, '联盟变化时序图.png', '联盟变化玩家', sort_by='alliance')
else:
    print(f"⚠️ 图1: 无有联盟变化+声望>{MIN_PRESTIGE_FOR_PLOT}的玩家")

if len(df_plot2) > 0:
    draw_timeline(df_plot2, '昵称变化时序图.png', '昵称变化玩家', sort_by='name')
else:
    print(f"⚠️ 图2: 无昵称变化+声望>{MIN_PRESTIGE_FOR_PLOT}的玩家")