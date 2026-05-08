# wjdr — 无尽冬日数据分析工具

## 项目概述
《无尽冬日》游戏的数据分析工具集，分析玩家行为（联盟变化、昵称变化、挖矿检测、战力对比、落单检测等）。

## 环境
- Python 3.14.3，依赖: `pandas`, `matplotlib`, `numpy`, `openpyxl`, `scikit-learn`, `PIL`
- 运行: `PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe script.py`

## 数据文件 (`data/`)
- 命名格式: `3957_MMDD.xlsx` 或 `3957_MMDD{a,b,c}.xlsx`
- 列: 账号, 大区, 炉子, 昵称, 联盟, 声望, 联盟名称, 坐标, 罩子, 援军数量, 兵线
- 豁免文件: `change_lm&name_protect.xlsx`（含 `账号`、`备注`）
- 数据整理: `data_prep.py` 可自动标准化文件名、修复声望进位bug、联盟列文本化

## 脚本功能

### `change_lm&name.py` — 昵称联盟变化分析 + 时序图
**输出 Excel 三表：** 变化总表（声望降序）→ 联盟变化排名 → 昵称变化排名
- 昵称/联盟取最早文件，声望取最新文件
- 变化格式：`a→b→c` 链条式
- 列: 账号, 昵称, 联盟, 声望, 昵称变化, 联盟变化
- 输出: `昵称联盟变化分析.xlsx` + `玩家变化时序图.png`

**🔧 配置区（脚本顶部）：**
- `DATA_FOLDER = 'data'`, `FILE_PATTERN = '3957_*.xlsx'`
- `READ_FILE_COUNT = 10` — 读取文件数
- `MIN_PRESTIGE_FOR_PLOT = 0`, `MAX_PRESTIGE = 20000`
- `MIN_ALLIANCE_CHANGES = 0`, `MIN_NAME_CHANGES = 0`, `MIN_TOTAL_CHANGES = 0`
- `TOP_N = 999`, `PROTECT_GREEN_ONLY = True`, `FILTER_LOW_POWER = True`, `MIN_POWER = 2000`

**图表布局规则：**
- 左侧：`(联盟)昵称` 右对齐，右侧：最新状态 + 声望条（严格等比）+ 声望数值
- 顶部表头：初始状态（右对齐）/ 声望 / 最新状态（左对齐）
- 底部三行：战力区间 / 规则1 / 规则2（pt 固定偏移）
- 标题、探查时间用华文行楷（STXingkai）；表头用 SimHei

### `Miner_find.py` — 疑似资源号检测
**🔧 配置区：**
- `READ_FILE_COUNT = 999`, `MIN_RECENT_MINING = 5`, `N_LATEST_DAYS = 5`
- `LATEST_THRESHOLD = 0.6`, `LATEST_MIN = 0.2`, `THRESHOLD = 0.49`
- `EXCLUDE_ALLIANCES = ['SSS', '999']`, `MAX_PRESTIGE = 1500`, `MAX_LEVEL = 25`

**输出：** `N_miner.xlsx`（绿色=近期活跃，蓝色=退游矿工）
- 列: 账号, 昵称, 联盟, 声望, 坐标, 罩子, 历史挖矿天数, 最新挖矿天数

### `map_view.py` — 菱形地图可视化 + 聚落识别
**输出：** `Map/地图_MMDD.png`
- 菱形坐标转换：`dx = gx - gy`, `dy = gx + gy`
- 聚落检测：迭代剔除离群点（半径≤30，最少20人），360°扫描找零遮挡标签位置
- 矿工高亮：双色竖切（左半原色 + 右半黑色）
- 地形分区：荒原(1200)、雪原(500)、沃土(200) 菱形环，以太阳城(597,597)为中心
- 附图（Nature-journal 风格）：左上玩家人数时序、左下炉子等级条形图、右下声望直方图+正态拟合
- DPI=1200，玩家边长=2

### `settlement_detail.py` — 聚落细节切块图
**输出：** `Map/聚落/{MMDD}/{联盟}_{人数}人.png`
- 对每个 ≥50 人的聚落按日切块放大，标注每个玩家名字
- 聚落检测：迭代剔除离群点（半径≤20），名字黑色，非聚落路过者灰色
- 玩家方块带灰色网格线（`#cccccc`, 0.3pt）以便区分毗邻玩家
- 名字长度检测：中文算2、英文算1，超10分两行显示
- 颜色：联盟映射优先，未映射联盟自动分配 AUTO_PALETTE 配色

**🔧 配置区：**
- `MIN_SETTLEMENT = 50` — 最少人数
- `VIEWPORT_PADDING = 0.35` — 包围盒外扩比例
- `VIEWPORT_SCALE = 1.0` — 视口放大倍数
- `NAME_FONTSIZE = 5`
- `DPI = 1200`

### 其他脚本
- `compare_zl_loss.py` — 战力损失对比
- `data_prep.py` — 数据整理（`DRY_RUN = False` 预览）
- `outlier_find.py` — 落单检测（DBSCAN 聚类）

## 项目特定规则（全局没有的）
- 图表配色: `RED_ALLIANCES=['FFF','666','OoO','SSR']→红`, `GREEN_ALLIANCES=['999','ann','xxx']→绿`, `BLUE_ALLIANCES=['KFC']→蓝`, 其他黑/灰
- 标注规则：联盟变化线上方标新名（联盟色），昵称变化线下方标新名（灰色）
- 底部定位用 `xytext=(0, -N)` + `offset points`，**不准用 transAxes**
- 昵称完整显示，边距自适应
- 代码注释极少（只写"为什么"）

## git
- 结果文件（`*.png`、`*.xlsx`）不提交，`.claude/` 不提交