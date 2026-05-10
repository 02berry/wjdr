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
- **声望基础单位为万**（数据中 10000 = 1亿）
- **注意：** Excel 列名叫"声望"，但游戏内玩家叫"战力"，代码中变量/输出统一用"声望"避免混淆

## 脚本功能

### `throne_photo.py` — 王城合照图
**输出：** `Map/{MMDD}/{MMDD}_王城合照.png`（DPI=1600，figsize=10×10）

#### 整体框架
- 读取 `data/` 最新 xlsx → 解析坐标 → 筛选视野 → 玩家着色 → 绘制地形+建筑 → 渲染玩家方块 → 头像叠加 → 太阳城建筑 → 标题/日期/右上文案 → 图例 → 输出

#### 坐标模型
```python
to_display(gx, gy) = (gx - gy, gx + gy)    # 45° 旋转
zone_diamond(cx, cy, half) → 菱形四顶点     # 在游戏坐标画正方形→菱形
```
- 玩家坐标 (gx,gy) 指向所占 2×2 块的**底部格子**
- 玩家菱形四顶点：`(gx,gy)→(gx+2,gy)→(gx+2,gy+2)→(gx,gy+2)`
- 地图几何中心：`MAP_CX, MAP_CY = 600, 600`
- dx = gx - gy, dy = gx + gy（菱形显示坐标）
- 太阳城建筑：游戏坐标 (597,597)~(603,603)，中心 (600,600)
- 王域：中心 (600,600)，half=6，范围 (594,594)~(606,606)
- 王域四角炮台：2×2方块，圆环相切，方向(中圈内切)

#### 视图范围
- 灰烬 half = `ASH_HALF`（默认 14），范围 (586,586)~(614,614)，最低点 586 不动，最高点 614，视口 = `2 × 2 × ASH_HALF × ASH_VIEWPORT_MULT`（显示坐标单位）
- 视野取 dx/dy 落在 `[center-half, center+half]` 内的玩家
- `ASH_VIEWPORT_MULT = 1.20`：超出灰烬边界 20%

#### 渲染分层（zorder）
```
zorder 0:  地形（沃土→废墟→灰烬→王域）
zorder 1:  建筑废墟方块
zorder 2:  玩家菱形方块（普通色/渐变）
zorder 2.1: 火焰晶金边
zorder 2.2: 破亿金色方块
zorder 2.5: 炮台圆环
zorder 3:  玩家名字
zorder 3.5: 02berry 头像
zorder 4:  太阳城建筑
zorder 5:  太阳城文字
```

#### 玩家着色逻辑（prestige 声望，单位=万）
```
is_gold = prestige ≥ GOLD_THRESHOLD（10000=1亿）
  gold → 金色方块，gold_norm = min((prestige-10000)/5000, 1.0)
         颜色从 #DAA520(#daa520) 渐变到 #FFD700(#ffd700)
elif prestige ≤ PRESTIGE_LOW（3000=3千万）
  → 最小联盟色（_soften(base, 0.90)，几乎白色）
else（3000~10000）
  → norm = (prestige - 3000) / 7000，以 base 为基准渐变加深
```
- `_soften(color, amount)`：向白色混合 amount 比例
- 火焰晶（炉子以"火"开头）：金边 `#FFE033` 0.6pt + 等级数字 `#FFF44F` 标注在 `(dx, dy+3.0)`
- **破亿玩家不叠加火晶标注**（两个标注不会同时出现）

#### 地形配置
- 沃土 half=149.5（`#e8f5e9`），废墟 half=48.5（`#e0e0e0`），灰烬 half=ASH_HALF（`#ffcdd2`）
- 王域 `#fff3cd`，half=6.0，中心 (600,600)
- 炮台：`(594,594),(604,594),(604,604),(594,604)` 四角，2×2 方块底色 `#ffe082` + 内切圆 `#ffb300`

#### 建筑废墟
- BUILDINGS 列表中包含 1 个太阳城 + 4 个要塞 + 12 个堡垒
- 要塞和堡垒画废墟方块：RUINS_HALF=29.5，仅画在视野内的
- 太阳城单独画：`#e8dcc8` 填充 + `#8b7355` 边线

#### 标题 / 日期 / 右上文案
- 标题：`TITLE`（默认"3957第三次王战大合照"）在 `(TITLE_X, TITLE_Y) = (0.02, 0.97)` 左对齐
- 日期：通过 `fig.canvas.draw()` 量取标题实际渲染宽度，右端对齐标题右端，y=0.932
- 右上文案：`RIGHT_TEXT` 在 `(0.98, 0.97)` 右对齐顶部，支持 `\n` 换行

#### 太阳城标注
- 中心坐标 `to_display(600, 600) = (0, 1200)`，文字 `SUN_CITY_LABEL`，字号 16 加粗

#### 02berry 头像
- 文件：`AVATAR_FILE`（默认 `纯白头像.png`），1254×1254
- 坐标：`AVATAR_COORDS = (606, 605)`
- 加载后 imshow 铺在菱形包围盒 `extent=[bgx-bgy-2, bgx-bgy+2, bgx+bgy, bgx+bgy+4]`
- 用 `ax.imshow()` + `set_clip_path()` 裁切成菱形，不遮挡相邻玩家
- 坐标在太阳城右上方

#### 图例（手动布局，底部靠右）
**算法：**
1. `sorted_ali` = 联盟名按人数升序（最少人在底部，最多人在顶部）
2. 临时测量第一行（最多人，sorted_ali[-1]）的联盟名渲染宽度 `name_r`
3. `cx = LX_N + name_r + 0.025` 确定人数列居中位置（0.025 = 4空格宽度）
4. `shift = cx - LX_C`，总位移 = `shift × 1.4`（初始1倍 + 额外0.4倍）
5. 每条：色条 `0.73→总位移后` + 联盟名左对齐 `0.78→总位移后` + 人数居中于 `cx→总位移后`

**关键坐标（初始，位移前）：**
- `LX_C = 0.73`（色条起点），色条宽 0.025
- `LX_N = 0.78`（联盟名起点）
- `LY0 = 0.02`（底部），`LY_STEP = 0.025`（行间距）
- 人数列使用 `ha='center'`，所有行对齐到同一 `cx`

#### 字体与编码
- SimSun（宋体）+ Times New Roman，`font.family = 'serif'`
- 中文字体注册：`simsun.ttc`、`times.ttf`、`msyh.ttc`、`simhei.ttf`（fallback）
- `axes.unicode_minus = False`

**🔧 完整配置区：**
| 变量 | 默认值 | 说明 |
|------|--------|------|
| `DATA_FOLDER` | `'data'` | 数据文件夹 |
| `MAP_FOLDER` | `'Map'` | 输出文件夹 |
| `DPI` | `1600` | 输出 DPI |
| `NAME_FONTSIZE` | `5` | 玩家名字字号 |
| `ASH_HALF` | `13.5` | 灰烬半宽（游戏坐标单位） |
| `ASH_VIEWPORT_MULT` | `1.20` | 视口放大系数 |
| `KINGDOM_HALF` | `6.0` | 王域半宽 |
| `RUINS_HALF` | `29.5` | 废墟方块半宽 |
| `GOLD_THRESHOLD` | `10000` | 破亿金色阈值（声望·万） |
| `PRESTIGE_LOW` | `3000` | 最浅档阈值 |
| `PRESTIGE_RANGE` | `7000` | 渐变范围（3000→10000） |
| `TITLE` | `'3957第三次王战大合照'` | 左上标题 |
| `TITLE_FONTSIZE` | `16` | 标题字号 |
| `TITLE_X, TITLE_Y` | `0.02, 0.97` | 标题位置（axes坐标） |
| `DATE_FONTSIZE` | `13` | 日期字号 |
| `RIGHT_TEXT` | 见代码 | 右上文案（支持 `\n`） |
| `RIGHT_FONTSIZE` | `11` | 右上文案字号 |
| `RIGHT_TEXT_X, RIGHT_TEXT_Y` | `0.98, 0.97` | 右上文案位置 |
| `SUN_CITY_LABEL` | `'太阳城'` | 太阳城标注文字 |
| `SUN_CITY_FONTSIZE` | `16` | 太阳城字号 |
| `AVATAR_FILE` | `'纯白头像.png'` | 头像图片路径 |
| `AVATAR_COORDS` | `(606, 605)` | 头像所在游戏坐标 |
| `BG_COLOR` | `'#f5f2ed'` | 背景色 |
| `OTHER_COLOR` | `'#ddd0bf'` | 未映射联盟默认色 |

---

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

### `map_view.py` — 菱形地图可视化 + 聚落识别 + 旗子领地 + 熊坑
**输出：** `Map/地图_MMDD.png`（1200DPI, 12×12英寸）

#### 功能
- 全地图渲染：同心菱形环地形（荒原→雪原→沃土→废墟→灰烬）+ 王域 + 炮台
- 前10联盟着色（联盟色 + 声望深浅），非前10黑色调（黑色=全服最高声望，依次减淡）
- 旗子/总部领地：格子级联合边界（重叠处无线，线宽0.15pt）
- 要塞/堡垒脚下废墟地形（`#f0f0f0`，覆盖原本雪原/沃土）
- 聚落标注 + 熊坑图片渲染
- 火晶玩家金边（0.15pt）
- 附图（Nature风格）：左上玩家人数时序、左下炉子等级条形图、右下声望直方图+正态拟合
- **无联盟 `-` 不占前10名额**

#### 坐标模型
```python
to_display(gx, gy) = (gx - gy, gx + gy)    # 45° 旋转 + √2 缩放
zone_diamond(cx, cy, half) → 菱形四顶点
```
- 玩家坐标 (gx,gy) 指向所占 2×2 块的**底部格子**
- 玩家菱形顶点：`(gx,gy)→(gx+2,gy)→(gx+2,gy+2)→(gx,gy+2)`
- 地图几何中心：`MAP_CX, MAP_CY = 600, 600`
- **旗子圆半径 = 1/√2**（内切于1×1格子），**HQ圆半径 = 3/√2**（内切于3×3块）

#### 渲染分层
```
zorder 0:   地形（荒原→雪原→沃土→废墟→灰烬→王域）
zorder 1:   要塞/堡垒废墟方块
zorder 1.5: 旗子领地填充
zorder 1.6: 领地联合边界
zorder 2:   玩家菱形块
zorder 2.1: 火晶金边
zorder 3:   建筑（太阳城、要塞、堡垒）
zorder 5:   旗子/HQ标记 + 聚落标注
zorder 6:   熊坑图片
```

#### 着色规则
- 前10联盟：`联盟色 + 声望深浅`（≤3000→90%白，3000~10000→渐变至本色，>10000→`_darken(base, 0.3)`）
- 非前10：黑色调，`ratio^1.2` 曲线（`ratio = 声望/全服最高声望`），低声望更浅
- 火晶玩家：`#FFE033` 金边 0.15pt

#### 🔧 配置区
- `TOP_N = 10` — 着色联盟数
- `NEUTRAL = '#000000'` — 非前10基础色
- `SETTLEMENT_MIN = 20`, `MAX_RADIUS = 30`
- `RUINS_HALF = 29.5`, `RUINS_COLOR = '#f0f0f0'`
- `ARC_CFG: flag {size:1, territory:7}, Headquarters {size:3, territory:15}`
- 地形颜色见 `ZONES` 列表
- 建筑坐标见 `BUILDINGS` 列表

### `settlement_detail.py` — 聚落细节切块图
**输出：** `Map/{MMDD}/{MMDD}_{联盟}.png`（1600DPI, 10×10英寸）
- 对每个 ≥50 人的聚落按日切块放大，标注每个玩家名字
- **坐标模型：** 同 map_view，玩家菱形占 (gx,gy)~(gx+2,gy+2)
- 聚落检测：迭代剔除离群点（半径≤20），名字黑色，非聚落路过者灰色
- 玩家方块带灰色网格线（`#cccccc`, 0.3pt）以便区分毗邻玩家
- 名字长度检测：中文算2、英文算1，超10分两行显示
- 名字标注在菱形中心 `(dx, dy+2)`
- 颜色：联盟映射优先，未映射联盟自动分配 AUTO_PALETTE 配色
- 火晶玩家网格线细红（火2加粗至0.8），声望过亿玩家金色方块+联盟色内切圆
- 颜色柔和化（`_soften` 向白色混合45%），暖白背景 `#f5f2ed`
- **熊坑检测：** 扫描 3×3 空格区域，坐标指向底部格子，曼哈顿距离选最优，菱形 `to_display(bx,by)→(bx+3,by)→(bx+3,by+3)→(bx,by+3)`

**🔧 配置区：**
- `MIN_SETTLEMENT = 50` — 最少人数
- `VIEWPORT_PADDING = 0.35` — 包围盒外扩比例
- `VIEWPORT_SCALE = 1.0` — 视口放大倍数
- `NAME_FONTSIZE = 5`
- `DPI = 1600`

### 其他脚本
- `compare_zl_loss.py` — 战力损失对比
- `data_prep.py` — 数据整理（`DRY_RUN = False` 预览）
- `outlier_find.py` — 落单检测（DBSCAN 聚类）

## 坐标模型（自然语言版）

地图渲染用 `to_display(gx, gy) = (gx - gy, gx + gy)` 把游戏坐标转成显示坐标。这等价于：
1. **把游戏格子逆时针旋转45度**
2. **边长放大√2倍**（所以游戏里边长1的正方形，旋转后菱形边长为√2）

换算关系：
- 游戏坐标 1×1 正方形（边长1）→ 显示坐标菱形（内切圆半径 = 0.5×√2 = 1/√2 ≈ 0.707）
- 游戏坐标 2×2 玩家块 → 显示坐标菱形（内切圆半径 = 2/√2 ≈ 1.414，即 to_display 后宽高各 2 单位）
- 换言之：用户说"一格的正方形"，对应显示坐标里"内切圆半径 0.707 的菱形"

## 项目特定规则（全局没有的）
- **改地图不动聚落**：改 `map_view.py` 时不要动 `settlement_detail.py`，两个脚本独立
- **不生成自检图**：改完直接运行，用户自己检查效果，不要生成低分辨率缩略图/测试图来自我验证
- 图表配色: `RED_ALLIANCES=['FFF','666','OoO','SSR']→红`, `GREEN_ALLIANCES=['999','ann','xxx']→绿`, `BLUE_ALLIANCES=['KFC']→蓝`, 其他黑/灰
- 标注规则：联盟变化线上方标新名（联盟色），昵称变化线下方标新名（灰色）
- 底部定位用 `xytext=(0, -N)` + `offset points`，**不准用 transAxes**
- 昵称完整显示，边距自适应
- 代码注释极少（只写"为什么"）

## git
- 结果文件（`*.png`、`*.xlsx`）不提交，`.claude/` 不提交
