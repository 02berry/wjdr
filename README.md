# 无尽冬日 数据分析工具

## 功能介绍

- **`data_prep.py`**：数据整理（文件名标准化、战力进位修复、联盟文本化），放新文件后先跑这个
- **`Miner_find.py`**：高频挖矿小号识别，历史+最新双维度筛选，近期活跃标绿、退游矿工标蓝
- **`change_lm&name.py`**：联盟/昵称变化追踪 + 时序图，三工作表输出（总表/联盟排名/昵称排名）
- **`map_view.py`**：菱形地图可视化，地形分区 + 聚落识别 + 矿工高亮 + 统计副图（Nature 风格）
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
- `Map/Architecture.xlsx`：建筑坐标配置（太阳城、要塞）

## 配置说明

修改各py文件顶部的配置区，关键参数：

| 脚本 | 关键配置 |
|------|----------|
| `data_prep.py` | `DRY_RUN`, `PRESTIGE_HIGH`, `PRESTIGE_BUG_MAX` |
| `Miner_find.py` | `READ_FILE_COUNT`, `N_LATEST_DAYS`, `LATEST_THRESHOLD`, `LATEST_MIN`, `THRESHOLD`, `EXCLUDE_ALLIANCES`, `MAX_PRESTIGE` |
| `change_lm&name.py` | `READ_FILE_COUNT`, `MAX_PRESTIGE`, `PROTECT_GREEN_ONLY`, `FILTER_LOW_POWER`, `MIN_POWER` |
| `map_view.py` | `S`, `PLAYER_SIZE`, `DPI`, `SETTLEMENT_MIN`, `MAX_RADIUS`, `ZONES`, `LABEL_OFFSET` |
