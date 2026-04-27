import pandas as pd
import re
import numpy as np
from sklearn.cluster import DBSCAN

# ==================== 1. 读取最新数据 ====================
df = pd.read_excel("3957_420.xlsx", sheet_name="全区")

print(f"总玩家数: {len(df)}")

# ==================== 2. 数据预处理 ====================
df["声望"] = pd.to_numeric(df["声望"], errors="coerce")
df["等级"] = df["炉子"].str.extract(r"lv\.(\d+)").astype(float)


def extract_xy(coord):
    if pd.isna(coord):
        return None, None
    match = re.search(r"\((\d+),\s*(\d+)\)", str(coord))
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None


df["X"], df["Y"] = zip(*df["坐标"].apply(extract_xy))
df = df[df["X"].notna() & df["Y"].notna()].copy()
print(f"有坐标的玩家数: {len(df)}")

# ==================== 3. 筛选大联盟 ====================
league_counts = df.groupby("联盟").size().reset_index(name="联盟人数")
big_leagues = league_counts[league_counts["联盟人数"] >= 20]["联盟"].tolist()

print(f"联盟总数: {len(league_counts)}")
print(f"人数≥20的大联盟数量: {len(big_leagues)}")

df_big = df[df["联盟"].isin(big_leagues)].copy()
print(f"大联盟玩家总数: {len(df_big)}")


# ==================== 4. 聚类找主聚落 ====================
def find_main_cluster_center(group, eps=150, min_samples=5):
    coords = group[["X", "Y"]].values

    if len(coords) < min_samples:
        return None, None, None

    clustering = DBSCAN(eps=eps, min_samples=min_samples).fit(coords)
    labels = clustering.labels_

    unique_labels, counts = np.unique(labels[labels >= 0], return_counts=True)

    if len(unique_labels) == 0:
        return None, None, None

    main_label = unique_labels[np.argmax(counts)]
    main_cluster = group[labels == main_label]

    center_x = main_cluster["X"].median()
    center_y = main_cluster["Y"].median()
    cluster_size = len(main_cluster)

    return center_x, center_y, cluster_size


def get_outliers_v2(group):
    """找脱离者：偏离主聚落 + 声望<1500 + 等级<20"""
    if len(group) < 20:
        return pd.DataFrame()

    center_x, center_y, cluster_size = find_main_cluster_center(group, eps=150, min_samples=5)

    if center_x is None:
        return pd.DataFrame()

    group["距离_X"] = abs(group["X"] - center_x)
    group["距离_Y"] = abs(group["Y"] - center_y)

    # 条件1：偏离超过200
    mask_distance = (group["距离_X"] > 200) | (group["距离_Y"] > 200)
    # 条件2：声望 < 1500
    mask_prestige = group["声望"] < 1500
    # 条件3：等级 < 20
    mask_level = group["等级"] < 20

    # 三个条件同时满足
    outliers = group[mask_distance & mask_prestige & mask_level].copy()

    if len(outliers) > 0:
        outliers["主聚落中心_X"] = center_x
        outliers["主聚落中心_Y"] = center_y
        outliers["主聚落人数"] = cluster_size
        outliers["联盟总人数"] = len(group)
        outliers["偏离_X"] = outliers["距离_X"]
        outliers["偏离_Y"] = outliers["距离_Y"]
        outliers["偏离描述"] = outliers.apply(
            lambda r: f"距主聚落({center_x:.0f},{center_y:.0f}) X偏差{r['距离_X']:.0f} Y偏差{r['距离_Y']:.0f}",
            axis=1
        )

    return outliers


# ==================== 5. 遍历所有大联盟 ====================
outliers_list = []
cluster_info_list = []

for league, group in df_big.groupby("联盟"):
    center_x, center_y, cluster_size = find_main_cluster_center(group, eps=150, min_samples=5)

    if center_x is not None:
        cluster_info_list.append({
            "联盟": league,
            "联盟总人数": len(group),
            "主聚落人数": cluster_size,
            "主聚落占比": f"{cluster_size / len(group) * 100:.1f}%",
            "主聚落中心_X": f"{center_x:.0f}",
            "主聚落中心_Y": f"{center_y:.0f}"
        })

    outliers = get_outliers_v2(group)
    if len(outliers) > 0:
        outliers_list.append(outliers)

# ==================== 6. 输出结果 ====================
if cluster_info_list:
    cluster_df = pd.DataFrame(cluster_info_list)
    cluster_df = cluster_df.sort_values("主聚落人数", ascending=False)
    print("\n各联盟主聚落信息:")
    print(cluster_df.to_string(index=False))

if outliers_list:
    outliers_df = pd.concat(outliers_list, ignore_index=True)

    output_cols = [
        "账号", "昵称", "炉子", "等级", "联盟", "声望", "坐标",
        "主聚落中心_X", "主聚落中心_Y", "主聚落人数", "联盟总人数",
        "偏离_X", "偏离_Y", "偏离描述"
    ]
    result = outliers_df[output_cols].copy()
    result = result.sort_values(["联盟", "偏离_X", "偏离_Y"], ascending=[True, False, False])

    print(f"\n找到符合条件的脱离者数量: {len(result)}")
    print("（条件：偏离主聚落 >200，声望 <1500，等级 <20）")

    league_outlier_counts = result.groupby("联盟").size().reset_index(name="脱离人数")
    league_outlier_counts = league_outlier_counts.sort_values("脱离人数", ascending=False)
    print("\n各联盟脱离人数统计:")
    print(league_outlier_counts.to_string(index=False))
else:
    result = pd.DataFrame()
    print("\n没有找到符合条件的脱离者")

# ==================== 7. 保存 ====================
with pd.ExcelWriter("脱离集体的玩家_聚类版.xlsx", engine="openpyxl") as writer:
    if len(result) > 0:
        result.to_excel(writer, sheet_name="脱离者名单", index=False)
    else:
        pd.DataFrame({"提示": ["没有找到符合条件的脱离者"]}).to_excel(writer, sheet_name="脱离者名单", index=False)

    if cluster_info_list:
        cluster_df.to_excel(writer, sheet_name="联盟聚落信息", index=False)

    if len(result) > 0:
        summary = result.groupby("联盟").agg(
            脱离人数=("账号", "count"),
            主聚落人数=("主聚落人数", "first"),
            联盟总人数=("联盟总人数", "first"),
            平均偏离X=("偏离_X", "mean"),
            平均偏离Y=("偏离_Y", "mean")
        ).reset_index().sort_values("脱离人数", ascending=False)
        summary.to_excel(writer, sheet_name="联盟汇总", index=False)

print("\n结果已保存: 脱离集体的玩家_聚类版.xlsx")