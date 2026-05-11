# 无尽冬日 数据分析工具

《无尽冬日》（Whiteout Survival）第 3957 王国的数据分析工具集，涵盖玩家行为追踪、地图可视化、资源号检测等功能。

## 功能展示

### 全地图渲染

![全地图](image/地图_0510.png)

同心菱形环地形（荒原→雪原→沃土→废墟→灰烬→王域），前 10 联盟按声望深浅着色，非前 10 黑色调渐变。旗子/总部领地联合边界，聚落标注 + 熊坑，附带人数/炉子等级/声望分布附图。

### 王城合照

![王城合照](image/0509_王城合照.png)

灰烬区域放大渲染，玩家按声望着色（过亿金色、火晶金边），支持头像叠加、太阳城建筑标注、图例自动布局。

### 聚落细节

| FFF 联盟 | 999 联盟 | KFC 联盟 |
|:---:|:---:|:---:|
| ![FFF](image/0510_FFF.png) | ![999](image/0510_999.png) | ![KFC](image/0510_KFC.png) |

对每个聚落单独切块放大，标注玩家名称、旗子领地、熊坑，灰色网格线区分毗邻玩家。

## 脚本一览

| 脚本 | 功能 |
|------|------|
| `data_prep.py` | 数据整理：文件名标准化、声望进位修复、联盟列文本化 |
| `Miner_find.py` | 疑似资源号检测：历史+最新双维度挖矿频率分析 |
| `change_lm&name.py` | 联盟/昵称变化追踪 + 时序图，三工作表输出 |
| `map_view.py` | 全地图可视化 + 聚落识别 + 旗子领地 + 熊坑 |
| `throne_photo.py` | 王城合照图，灰烬区域放大 + 头像叠加 |
| `settlement_detail.py` | 聚落细节切块图，单联盟放大 + 玩家名标注 |

## 快速开始

1. 将游戏导出的 Excel 放入 `data/` 目录（命名 `3957_MMDD.xlsx`）
2. 运行 `data_prep.py` 整理数据
3. 运行目标脚本：`PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe script.py`
4. 地图输出到 `Map/`，分析结果输出到当前目录

## 数据文件

| 文件 | 说明 |
|------|------|
| `data/3957_MMDD.xlsx` | 游戏数据快照 |
| `data/change_lm&name_protect.xlsx` | 变化追踪豁免名单 |
| `Map/league_mapping.xlsx` | 联盟→颜色映射表 |
| `Map/arc.xlsx` | 旗子/总部坐标表 |

## 关键配置

各脚本顶部配置区可调参数：

| 脚本 | 关键参数 |
|------|----------|
| `data_prep.py` | `DRY_RUN`, `PRESTIGE_HIGH`, `PRESTIGE_BUG_MAX` |
| `Miner_find.py` | `READ_FILE_COUNT`, `THRESHOLD`, `MAX_PRESTIGE`, `EXCLUDE_ALLIANCES` |
| `change_lm&name.py` | `READ_FILE_COUNT`, `MAX_PRESTIGE`, `PROTECT_GREEN_ONLY`, `MIN_POWER` |
| `map_view.py` | `TOP_N=10`, `NEUTRAL=#000000`, `SETTLEMENT_MIN=20`, `ARC_CFG` |
| `throne_photo.py` | `ASH_HALF=13.5`, `TITLE`, `GOLD_THRESHOLD=10000` |
| `settlement_detail.py` | `MIN_SETTLEMENT=50`, `VIEWPORT_PADDING=0.35`, `DPI=1600` |

---

<span style="color:red"><b>⚠ 数据依赖说明：</b>本工具集需要游戏内导出数据才能运行，需要全区玩家坐标，推荐白猫工具。不同王国迁移需要自行适配脚本，不能直接照搬。</span>
