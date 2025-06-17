import os
import pandas as pd
import json


# === 設定 ===
log_dir = "log"
excel_dir = os.path.join("log")
os.makedirs(excel_dir, exist_ok=True)
output_excel_path = os.path.join(excel_dir, "SECTNO_drug_qty.xlsx")

# === 解析 log 每筆 ===
def extract_preferred_sectno(med_log_list):
    for entry in med_log_list:
        sectno = entry.get("所需欄位", {}).get("SECTNO")
        if sectno and sectno != "None":
            return sectno
    return "未知科別"

# === 累積每天的資料成 {科別: dataframe} 結構 ===
dept_data_by_day = {}

for filename in sorted(os.listdir(log_dir)):
    if filename.endswith(".txt"):
        date_str = filename.replace(".txt", "")  # 20250617

        daily_records = []
        with open(os.path.join(log_dir, filename), "r", encoding="utf-8") as f:
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
                            daily_records.append({
                                "藥品名稱": drug_name,
                                "SECTNO": sectno,
                                "TXN_QTY": txn_qty
                            })

                except json.JSONDecodeError:
                    continue

        # 按科別分組當日藥品用量
        df_day = pd.DataFrame(daily_records)
        if df_day.empty:
            continue

        for sectno, group in df_day.groupby("SECTNO"):
            summary = group.groupby("藥品名稱")["TXN_QTY"].sum().reset_index()
            summary.columns = ["藥品名稱", date_str]

            if sectno not in dept_data_by_day:
                dept_data_by_day[sectno] = summary
            else:
                # 先刪除舊日期欄（若存在）
                if date_str in dept_data_by_day[sectno].columns:
                    dept_data_by_day[sectno] = dept_data_by_day[sectno].drop(columns=[date_str])

                # 再合併新欄位
                dept_data_by_day[sectno] = pd.merge(
                    dept_data_by_day[sectno], summary, on="藥品名稱", how="outer"
                ).fillna(0)

# === 排序與寫入 Excel ===
with pd.ExcelWriter(output_excel_path, engine="openpyxl", mode="w") as writer:
    for sect, df_sheet in dept_data_by_day.items():
        df_sheet = df_sheet.sort_values(by="藥品名稱")
        df_sheet.to_excel(writer, sheet_name=sect[:31], index=False)

print(f"✅ 已根據 txt 檔名寫入 Excel：{output_excel_path}")
