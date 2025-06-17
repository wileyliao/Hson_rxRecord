import os
import json
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import matplotlib

# 字型設定
matplotlib.rcParams['font.family'] = 'Microsoft JhengHei'
matplotlib.rcParams['axes.unicode_minus'] = False

# 圖片儲存資料夾
plot_dir = os.path.join("log", "plot")
os.makedirs(plot_dir, exist_ok=True)

# 批次讀取 log txt
log_dir = "log"
records = []

def extract_preferred_sectno(med_log_list):
    for entry in med_log_list:
        sectno = entry.get("所需欄位", {}).get("SECTNO")
        if sectno and sectno != "None":
            return sectno
    return "未知科別"

for filename in os.listdir(log_dir):
    if filename.endswith(".txt"):
        file_path = os.path.join(log_dir, filename)
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    log = json.loads(line)
                    med_list = log.get("藥品列表", [])
                    # 替換 for med in med_list: 區塊
                    for med in med_list:
                        drug_name = med.get("藥品名稱", "")
                        med_log_list = med.get("log", [])
                        sectno = extract_preferred_sectno(med_log_list)

                        # ➤ 每個藥只計一次 TXN_QTY
                        txn_qty = None
                        for rule in med_log_list:
                            fields = rule.get("所需欄位", {})
                            if txn_qty is None and "TXN_QTY" in fields and str(fields["TXN_QTY"]).isdigit():
                                txn_qty = int(fields["TXN_QTY"])

                        if txn_qty is not None:
                            icd = next((r.get("所需欄位", {}).get("ICD_CODE") for r in med_log_list if
                                        "ICD_CODE" in r.get("所需欄位", {})), "None")
                            records.append({
                                "藥袋編號": log.get("藥袋編號", ""),
                                "NAME": drug_name,
                                "ICD_CODE": icd,
                                "SECTNO": sectno,
                                "TXN_QTY": txn_qty
                            })

                except json.JSONDecodeError:
                    continue

# 整理成 DataFrame
df = pd.DataFrame(records)
if df.empty:
    print("❌ 沒有找到任何 log 資料。")
    exit()

# [1] 藥品使用量前十名
top10 = df.groupby("NAME")["TXN_QTY"].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(10, 6))
top10.plot(kind="bar", color="skyblue")
plt.title("藥品使用量前十名（依總量）")
plt.ylabel("總量")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.grid(axis="y")
plt.savefig(os.path.join(plot_dir, "top10_drugs.png"))
plt.show()

# [2] 科別藥品堆疊圖
pivot = df.groupby(["SECTNO", "NAME"])["TXN_QTY"].sum().unstack().fillna(0)
pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).head(10).index]
pivot.plot(kind="bar", stacked=True, figsize=(12, 6), colormap="tab20")
plt.title("科別藥品用量分布")
plt.ylabel("總使用量")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.grid(axis="y")
plt.savefig(os.path.join(plot_dir, "dept_drug_distribution.png"))
plt.show()

# [3] 診斷碼 → 藥品用量互動 bar chart（含 dropdown）
grouped = df.groupby(["ICD_CODE", "NAME"])["TXN_QTY"].sum().reset_index()
icd_list = grouped["ICD_CODE"].unique().tolist()

fig = go.Figure()


# 預設第一筆 ICD_CODE 當作初始化
initial_icd = icd_list[0]
init_data = grouped[grouped["ICD_CODE"] == initial_icd]

fig.add_trace(go.Bar(
    x=init_data["NAME"],
    y=init_data["TXN_QTY"],
    name=initial_icd
))

# 加入下拉選單
dropdown_buttons = []
for icd in icd_list:
    filtered = grouped[grouped["ICD_CODE"] == icd]
    dropdown_buttons.append(dict(
        label=icd,
        method="update",
        args=[{
            "x": [filtered["NAME"]],
            "y": [filtered["TXN_QTY"]],
            "type": "bar"
        }, {
            "title": f"診斷碼：{icd} 的用藥情形"
        }]
    ))

fig.update_layout(
    title=f"診斷碼：{initial_icd} 的用藥情形",
    xaxis_title="藥品名稱",
    yaxis_title="總使用量",
    updatemenus=[dict(
        active=0,
        buttons=dropdown_buttons,
        x=1.15,
        xanchor="left",
        y=1.15,
        yanchor="top"
    )],
    autosize=True
)

fig.show()  # 不儲存，只展示
