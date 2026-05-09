# 无尽冬日 数据分析工具

## 功能介绍

- **`data_prep.py`**：数据整理（文件名标准化、战力进位修复、联盟文本化），放新文件后先跑这个
- **`Miner_find.py`**：高频挖矿小号识别，历史+最新双维度筛选，近期活跃标绿、退游矿工标蓝
- **`change_lm&name.py`**：联盟/昵称变化追踪 + 时序图，三工作表输出（总表/联盟排名/昵称排名）
- **`throne_photo.py`**：王城合照图，灰烬区域放大，联盟颜色+声望深浅+火晶/破亿渲染
- **`settlement_detail.py`**：聚落细节切块图，按日分文件夹输出
- **`compare_zl_loss.py`**：声望战损对比，按联盟排名
- **`outlier_find.py`**：DBSCAN 聚类落单检测

## 使用方法

1. 将导出的Excel文件放入 `data` 文件夹
2. 运行 `data_prep.py` 整理数据（可选）
3. 运行分析脚本：`PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe script.py`
4. 结果输出到当前目录（地图输出到 `Map/` 目录）

## 数据文件

- `data/3957_MMDD.xlsx` 或 `3957_MMDD{a,b,c}.xlsx`（同一天多次快照）
- `data/change_lm&name_protect.xlsx`：变化追踪豁免名单
- `Map/league_mapping.xlsx`：联盟→颜色映射表

## 配置说明

修改各py文件顶部的配置区，关键参数：

| 脚本 | 关键配置 |
|------|----------|
| `data_prep.py` | `DRY_RUN`, `PRESTIGE_HIGH`, `PRESTIGE_BUG_MAX` |
| `Miner_find.py` | `READ_FILE_COUNT`, `N_LATEST_DAYS`, `LATEST_THRESHOLD`, `LATEST_MIN`, `THRESHOLD`, `EXCLUDE_ALLIANCES`, `MAX_PRESTIGE` |
| `change_lm&name.py` | `READ_FILE_COUNT`, `MAX_PRESTIGE`, `PROTECT_GREEN_ONLY`, `FILTER_LOW_POWER`, `MIN_POWER` |
| `map_view.py` | `S`, `PLAYER_SIZE`, `DPI`, `SETTLEMENT_MIN`, `MAX_RADIUS`, `ZONES`, `LABEL_OFFSET` |
| `throne_photo.py` | `ASH_HALF`, `ASH_VIEWPORT_MULT`, `TITLE`, `RIGHT_TEXT`, `AVATAR_FILE`, `AVATAR_COORDS`, `GOLD_THRESHOLD`, `PRESTIGE_LOW` |
| `settlement_detail.py` | `MIN_SETTLEMENT`, `VIEWPORT_PADDING`, `VIEWPORT_SCALE`, `NAME_FONTSIZE`, `DPI` |
