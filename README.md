# 无尽冬日 数据分析工具

## 功能介绍

- **`Miner_find.py`**：高频挖矿小号识别，多快照对比，支持备注标记、退游矿工标蓝、联盟排除、声望过滤
- **`change_lm&name.py`**：联盟/昵称变化追踪 + 时序图可视化
- **`compare_zl_loss.py`**：声望战损对比，按联盟排名
- **`outlier_find.py`**：DBSCAN 聚类落单检测

## 使用方法

1. 将导出的Excel文件放入 `data` 文件夹
2. 运行对应的py文件：`PYTHONIOENCODING=utf-8 .venv/Scripts/python.exe script.py`
3. 结果输出到当前目录

## 数据文件

- `data/3957_MMDD.xlsx` 或 `3957_MMDD{a,b,c}.xlsx`（同一天多次快照）
- `data/miner_remark.xlsx`：矿工备注（账号、备注1、备注2、备注3）
- `data/change_lm&name_protect.xlsx`：变化追踪豁免名单

## 配置说明

修改各py文件顶部的 `🔧 配置区`，关键参数：

| 脚本 | 关键配置 |
|------|----------|
| `Miner_find.py` | `READ_FILE_COUNT`, `MIN_RECENT_MINING`, `EXCLUDE_ALLIANCES`, `THRESHOLD`, `MAX_PRESTIGE`, `MAX_LEVEL` |
| `change_lm&name.py` | `READ_FILE_COUNT`, `MAX_PRESTIGE`, `PROTECT_GREEN_ONLY`, `FILTER_LOW_POWER`, `MIN_POWER` |
