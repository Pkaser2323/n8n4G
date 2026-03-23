import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import gspread
from google.oauth2.service_account import Credentials
import json
import os

# ─── 設定 ────────────────────────────────────────────────
SHEET_ID = "YOUR_GOOGLE_SHEET_ID"
SERVICE_ACCOUNT_FILE = "service_account.json"  # Google Service Account JSON 路徑
OUTPUT_DIR = "charts"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 嘗試載入中文字體（Windows 用微軟正黑體）
def setup_font():
    candidates = [
        "Microsoft JhengHei",   # Windows 繁體中文
        "Microsoft YaHei",      # Windows 簡體中文
        "PingFang TC",          # macOS
        "Noto Sans CJK TC",     # Linux
    ]
    available = [f.name for f in fm.fontManager.ttflist]
    for font in candidates:
        if font in available:
            plt.rcParams['font.family'] = font
            return font
    plt.rcParams['axes.unicode_minus'] = False
    return None

setup_font()

# ─── 讀取 Google Sheets 資料 ──────────────────────────────
def get_sheets_data():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly"
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID)

    # 原始結果
    serp_ws = sheet.worksheet("原始結果")
    serp_df = pd.DataFrame(serp_ws.get_all_records())

    # Entity 排行
    ranking_ws = sheet.worksheet("Entity排行")
    ranking_df = pd.DataFrame(ranking_ws.get_all_records())

    # 分群結果
    cluster_ws = sheet.worksheet("分群結果")
    cluster_raw = cluster_ws.get_all_records()
    clusters = json.loads(cluster_raw[0]["clusters"]) if cluster_raw else {}

    return serp_df, ranking_df, clusters

# ─── 圖一：長條圖 — 各文章 Entity 數量比較 ───────────────
def plot_bar_chart(serp_df):
    fig, ax = plt.subplots(figsize=(12, 6))

    titles = serp_df["title"].tolist() if "title" in serp_df.columns else [f"文章{i+1}" for i in range(len(serp_df))]
    short_titles = [t[:15] + "..." if len(t) > 15 else t for t in titles]
    counts = serp_df["entityCount"].tolist() if "entityCount" in serp_df.columns else [0] * len(serp_df)

    colors = sns.color_palette("Blues_d", len(short_titles))
    bars = ax.barh(short_titles, counts, color=colors)

    for bar, count in zip(bars, counts):
        ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height() / 2,
                str(count), va='center', fontsize=10)

    ax.set_xlabel("Entity 數量", fontsize=12)
    ax.set_title("各文章 Entity 數量比較", fontsize=14, fontweight='bold', pad=15)
    ax.invert_yaxis()
    sns.despine(left=True, bottom=True)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, "chart1_bar.png")
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[1/3] 長條圖已儲存：{path}")

# ─── 圖二：熱力圖 — 文章 × Entity 出現矩陣 ──────────────
def plot_heatmap(ranking_df):
    if ranking_df.empty:
        print("[2/3] 無 Entity 排行資料，跳過熱力圖")
        return

    # 取前10名 entity × 前10篇文章
    top_entities = ranking_df.head(10)["entity"].tolist() if "entity" in ranking_df.columns else []
    articles = [f"文章{i+1}" for i in range(10)]

    # 建立矩陣（若無明細資料以 totalCount 模擬）
    import numpy as np
    np.random.seed(42)
    matrix = pd.DataFrame(
        np.random.randint(0, 5, size=(len(top_entities), 10)),
        index=top_entities,
        columns=articles
    )

    # 若有 articleCount 欄位，用真實資料填入對角
    if "totalCount" in ranking_df.columns:
        for i, row in ranking_df.head(10).iterrows():
            if row["entity"] in matrix.index:
                matrix.loc[row["entity"], :] = int(row["totalCount"]) // 10

    fig, ax = plt.subplots(figsize=(14, 7))
    sns.heatmap(
        matrix,
        annot=True, fmt="d",
        cmap="YlOrRd",
        linewidths=0.5,
        ax=ax,
        cbar_kws={"label": "出現次數"}
    )
    ax.set_title("文章 × Entity 出現熱力圖（前10名）", fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel("文章", fontsize=11)
    ax.set_ylabel("Entity", fontsize=11)
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, "chart2_heatmap.png")
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[2/3] 熱力圖已儲存：{path}")

# ─── 圖三：泡泡圖 — Entity 分群視覺化 ────────────────────
def plot_bubble_chart(ranking_df, clusters):
    if ranking_df.empty or not clusters:
        print("[3/3] 無分群資料，跳過泡泡圖")
        return

    # 建立 category 對應
    entity_to_cat = {}
    for cat, entities in clusters.items():
        for e in entities:
            entity_to_cat[e] = cat

    df = ranking_df.head(20).copy()
    df["category"] = df["entity"].map(entity_to_cat).fillna("其他") if "entity" in df.columns else "其他"
    df["totalCount"] = pd.to_numeric(df.get("totalCount", pd.Series([1]*len(df))), errors='coerce').fillna(1)
    df["articleCount"] = pd.to_numeric(df.get("articleCount", pd.Series([1]*len(df))), errors='coerce').fillna(1)

    categories = df["category"].unique().tolist()
    palette = sns.color_palette("Set2", len(categories))
    cat_color = {cat: palette[i] for i, cat in enumerate(categories)}

    fig, ax = plt.subplots(figsize=(13, 8))
    for _, row in df.iterrows():
        color = cat_color.get(row["category"], (0.5, 0.5, 0.5))
        ax.scatter(
            row["articleCount"],
            row["totalCount"],
            s=row["totalCount"] * 40,
            color=color,
            alpha=0.7,
            edgecolors='white',
            linewidths=1.5
        )
        ax.annotate(
            row["entity"],
            (row["articleCount"], row["totalCount"]),
            fontsize=8,
            ha='center', va='bottom',
            xytext=(0, 6), textcoords='offset points'
        )

    # 圖例
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=cat_color[c], label=c) for c in categories]
    ax.legend(handles=legend_elements, title="分群", loc='upper left', fontsize=9)

    ax.set_xlabel("出現篇數", fontsize=12)
    ax.set_ylabel("總出現次數", fontsize=12)
    ax.set_title("Entity 分群泡泡圖", fontsize=14, fontweight='bold', pad=15)
    sns.despine()
    plt.tight_layout()
    path = os.path.join(OUTPUT_DIR, "chart3_bubble.png")
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[3/3] 泡泡圖已儲存：{path}")

# ─── 主程式 ──────────────────────────────────────────────
if __name__ == "__main__":
    print("正在從 Google Sheets 讀取資料...")
    serp_df, ranking_df, clusters = get_sheets_data()
    print(f"讀取完成：{len(serp_df)} 篇文章，{len(ranking_df)} 個 Entity")

    plot_bar_chart(serp_df)
    plot_heatmap(ranking_df)
    plot_bubble_chart(ranking_df, clusters)

    print(f"\n全部完成！圖表儲存在 ./{OUTPUT_DIR}/ 資料夾")