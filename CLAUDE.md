# wjdr — 无尽冬日数据分析工具

## 项目概述
《无尽冬日》游戏的数据分析工具集，分析玩家行为（联盟变化、昵称变化、挖矿检测、战力对比、落单检测等）。

## 环境
- Python 3.14.3 (`.venv` 虚拟环境)
- 依赖: `pandas`, `matplotlib`, `numpy`, `openpyxl`, `scikit-learn`, `PIL`
- 运行脚本用: `PYTHONIOENCODING=utf-8 ".venv/Scripts/python.exe" script.py`

## 数据文件 (`data/`)
- 命名格式: `3957_MMDD.xlsx` 或 `3957_MMDD{a,b,c}.xlsx`（同一天多次）
- 列: 账号, 大区, 炉子, 昵称, 联盟, 声望, 联盟名称, 坐标, 罩子, 援军数量, 兵线
- 豁免文件: `change_lm&name_protect.xlsx`（含 `账号`、`备注` 列）
- 矿工备注: `miner_remark.xlsx`（含 `账号`、`备注1`、`备注2`、`备注3` 列）

## 脚本功能

### `change_lm&name.py` — 昵称联盟变化分析 + 时序图
读取多个时间点的 xlsx 快照，追踪玩家联盟/昵称变化，输出 Excel + 单张合并时序图。

**关键配置区（脚本顶部）：**
- `DATA_FOLDER = 'data'`
- `FILE_PATTERN = '3957_*.xlsx'`
- `READ_FILE_COUNT = 10` — 读取文件数（最旧1个 + 最新N-1个）
- `MIN_PRESTIGE_FOR_PLOT = 0` — 图表筛选的最小声望
- `MAX_PRESTIGE = 20000` — 声望上限（超此值截断，防异常数据撑坏比例）
- `MIN_ALLIANCE_CHANGES = 0` / `MIN_NAME_CHANGES = 0` / `MIN_TOTAL_CHANGES = 0` — 变化次数阈值
- `TOP_N = 999` — 图表展示前 N 个玩家
- `PROTECT_GREEN_ONLY = True` — 保护友方：排除仅出现绿色联盟的玩家
- `FILTER_LOW_POWER = True` + `MIN_POWER = 2000` — 低战过滤：排除声望<阈值且无红色历史的玩家
- 豁免账号通过 EXEMPT_ACCOUNTS 列表编程添加（非外部文件）

**输出：** `昵称联盟变化分析.xlsx` + `玩家变化时序图.png`

**图表布局规则：**
- 单张图，按声望降序排列，含退游过滤（最新快照未出现者排除）
- 左侧：初始状态标签 `(联盟)昵称`，右对齐
- 右侧：最新状态标签 `(联盟)昵称`（左对齐）+ 声望条（严格等比于 max_prestige）+ 声望数值
- 顶部表头行：初始状态（右对齐）/ 声望（左对齐柱状图）/ 最新状态（左对齐）
- 声望条起始偏移按右侧最长标签动态计算（`bar_offset_mult`），防重叠
- 初始/最新数据点无内嵌标注，仅中间变化点标注联盟名（线上方）和昵称（线下方）
- 底部三行输出规则（pt 固定偏移）：战力区间 / 规则1（绿色过滤）/ 规则2（低战过滤）
- 标题、探查时间用华文行楷（STXingkai）；表头用 SimHei
- 图例一行无框，底部元素间距用 offset points（不受图高缩放影响）

### `Miner_find.py` — 疑似资源号检测
检测多个快照中持续在外采集的账号，输出 Excel。

**关键配置区（脚本顶部）：**
- `READ_FILE_COUNT = 10` — 读取最近N天的数据
- `MIN_RECENT_MINING = 5` — 最近N次无挖矿标蓝（退游矿工）
- `EXCLUDE_ALLIANCES = ['SSS', '999']` — 排除的联盟列表
- `THRESHOLD = 0.49` — 挖矿率阈值
- `MAX_PRESTIGE = 1500` — 声望上限（用最新快照声望判断）
- `MAX_LEVEL = 25` — 炉子等级上限

**输出：** `N_miner.xlsx`（N为分析次数），宋体14号居中对齐，无筛选
- 红色行 = 有备注的账号，蓝色行 = 退游矿工
- 列: 账号, 昵称, 联盟, 声望, 坐标, 罩子, 挖矿天数, 挖矿比例/%, 备注1, 备注2, 备注3
- 按挖矿率降序排列

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
- 联盟变化：线上方标注新联盟名（联盟色加粗 + 白色描边）
- 昵称变化：线下方标注新昵称（灰色加粗 + 白色描边）
- 初始/最新数据点不标注（左右侧标签已展示）
- 规则1/2 描述从 PROTECT_GREEN_ONLY / FILTER_LOW_POWER + MIN_POWER 读取

## 工作偏好
- **一次只改一个点**，改完跑图确认再下一个。不要批量改动。
- 底部元素用 `xytext=(0, -N)` + `textcoords='offset points'` 定位，**不准用 transAxes**（图高缩放导致间距失控）。
- 图表位置我用 pt 值报给你，直接调 pt 值，不要重新设计间距算法。
- 昵称完整显示不截断，边距按最长标签自适应。
- 代码注释极少（只写"为什么"），回复简短不用收尾总结。
- 不要做我没要求的事，不要"顺便"优化别的。
- 配置项放脚本顶部 `🔧 配置区`，让我自己改。
- 阶段完成后 git commit + push，message 用中文。

## git
- 生成的结果文件（`*.png`、`昵称联盟变化分析.xlsx` 等）不提交
- `.claude/` 不提交（已在 `.gitignore`）
