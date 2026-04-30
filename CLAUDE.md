# wjdr — 无尽冬日数据分析工具

## 项目概述
《无尽冬日》游戏的数据分析工具集，分析玩家行为（联盟变化、昵称变化、挖矿检测、战力对比、落单检测等）。

## 环境
- Python 3.14.3 (`.venv` 虚拟环境)
- 依赖: `pandas`, `matplotlib`, `numpy`, `openpyxl`, `scikit-learn`, `PIL`
- 运行脚本用: `PYTHONIOENCODING=utf-8 ".venv/Scripts/python.exe" script.py`

## 数据文件 (`data/`)
- 命名格式: `3957_MMDD.xlsx` 或 `3957_MMDD{a,b,c}.xlsx`（同一天多次）
- 列: 账号, 联盟, 炉子, 昵称, 坐标, 声望, 兵力, 采集队伍, 援助兵力, 驻防
- 豁免文件: `change_lm&name_protect.xlsx`（含 `账号`、`备注` 列）

## 脚本功能

### `change_lm&name.py` — 昵称联盟变化分析 + 时序图
读取多个时间点的 xlsx 快照，追踪玩家联盟/昵称变化，输出 Excel + 两张时序图。

**关键配置区（脚本顶部）：**
- `DATA_FOLDER = 'data'`
- `FILE_PATTERN = '3957_*.xlsx'`
- `READ_FILE_COUNT = 10` — 读取文件数（最旧1个 + 最新N-1个）
- `MIN_PRESTIGE_FOR_PLOT = 3000` — 图表筛选的最小声望
- `TOP_N = 40` — 图表展示前N个玩家

**输出：** `昵称联盟变化分析.xlsx` + `联盟变化时序图.png` + `昵称变化时序图.png`

### `Miner_find.py` — 疑似资源号检测
检测多个快照中持续在外采集的账号。

### `compare_zl_loss.py` — 战力损失对比
比较两个时间点的玩家战力变化。

### `outlier_find.py` — 落单检测
用 DBSCAN 聚类找出脱离联盟聚落的孤立玩家。

## 图表配色规则
```python
RED_ALLIANCES = ['FFF', '666', 'OoO', 'SSR']   → '#E74C3C' (红色)
GREEN_ALLIANCES = ['999', 'ann', 'xxx']          → '#27AE60' (绿色)
BLUE_ALLIANCES = ['KFC']                         → '#3498DB' (蓝色)
其他联盟 → '#000000' (黑色)
无联盟  → '#888888' (灰色)
```

## 图表标注规则（`draw_timeline` 函数）
- 联盟变化（如 `999→FFF`）: 多色分段着色，旧联盟用旧色、箭头灰色、新联盟用新色
- 昵称变化（如 `旧名→新名`）: 灰色
- 同时变化: `999→FFF | 旧名→新名` 格式，`|` 分隔
- 水平重叠自动避让：相邻标注间距小于阈值时交替上下放置
- 左侧起始信息带圆角背景框（联盟色描边）
- 右侧最新昵称带圆角背景框（联盟色描边），与左侧统一格式
- 右侧统计区：变化次数 + 声望条 + 声望值（千分位）
- 图例水平一行置于图表底部

## git
- 生成的结果文件（`*.png`、`昵称联盟变化分析.xlsx` 等）不提交
- `.claude/` 不提交（已在 `.gitignore`）
