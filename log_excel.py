import os
import pandas as pd
import json
from datetime import datetime

# === 設定 ===
log_dir = "log"
excel_path = os.path.join(log_dir, "SECTNO_drug_qty.xlsx")
today_str = datetime.now().strftime("%Y%m%d")

# === 取得藥袋第一個 SECTNO ===
def get_first_sectno(med_list):
    for med in med_list:
        for log_entry in med.get("log", []):
            sectno = log_entry.get("所需欄位", {}).get("SECTNO")
            if sectno and sectno != "None":
                return sectno
    return "未知科別"

# === 讀取原有 Excel 資料（如果存在） ===
dept_data_by_day = {}
if os.path.exists(excel_path):
    xls = pd.ExcelFile(excel_path)
    for sheet in xls.sheet_names:
        dept_data_by_day[sheet] = xls.parse(sheet)

# === 處理每個 log 檔案 ===
for filename in sorted(os.listdir(log_dir)):
    if not filename.endswith(".txt"):
        continue
    if filename.startswith(today_str):
        print(f"🛑 排除今日檔案：{filename}")
        continue

    date_str = filename.replace(".txt", "")
    print(f"🔍 正在處理：{filename}")

    daily_records_by_sectno = {}

    with open(os.path.join(log_dir, filename), "r", encoding="utf-8") as f:
        for line in f:
            try:
                log = json.loads(line)
                med_list = log.get("規則審查", [])
                if not med_list:
                    continue

                sectno = get_first_sectno(med_list)
                if not sectno:
                    continue

                for med in med_list:
                    name = med.get("藥品名稱", "")
                    txn_qty = None
                    for rule in med.get("log", []):
                        fields = rule.get("所需欄位", {})
                        if txn_qty is None and "TXN_QTY" in fields and str(fields["TXN_QTY"]).isdigit():
                            txn_qty = int(fields["TXN_QTY"])
                    if txn_qty is not None:
                        daily_records_by_sectno.setdefault(sectno, []).append({
                            "藥品名稱": name,
                            "TXN_QTY": txn_qty
                        })
            except Exception:
                continue

    if not daily_records_by_sectno:
        print(f"⚠️ {filename} 無有效資料，跳過")
        continue

    for sectno, records in daily_records_by_sectno.items():
        df_new = pd.DataFrame(records)
        if df_new.empty:
            continue

        summary = df_new.groupby("藥品名稱")["TXN_QTY"].sum().reset_index()
        summary.columns = ["藥品名稱", date_str]

        if sectno not in dept_data_by_day:
            dept_data_by_day[sectno] = summary
        else:
            old_df = dept_data_by_day[sectno]
            if date_str in old_df.columns:
                old_df = old_df.drop(columns=[date_str])
            dept_data_by_day[sectno] = pd.merge(old_df, summary, on="藥品名稱", how="outer").fillna(0)

    # 刪除已處理的 log
    os.remove(os.path.join(log_dir, filename))
    print(f"🗑 已刪除已處理的 log：{filename}")

# === 寫入 Excel ===
with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
    for sectno, df in dept_data_by_day.items():
        df = df.sort_values(by="藥品名稱")
        df.to_excel(writer, sheet_name=sectno[:31], index=False)

print(f"✅ 已寫入 Excel：{excel_path}")
