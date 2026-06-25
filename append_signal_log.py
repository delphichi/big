# -*- coding: utf-8 -*-
"""
每日訊號 Log (Daily Signal Log)
=====================================================================
每天收盤後,把 PE 監看「全部監看」分頁的當日推薦快照 append 一行到
data/daily_signal_log.csv。累積後可用 backtest_signals.py 接上未來報酬,
驗證「我們各種訊號標籤(⭐A級買點 / ⚠️便宜陷阱 …)的實際命中率」。

★ 為什麼要 append 而非覆蓋:
  原本每天覆蓋掉昨天的「訊號 vs 當天股價」→ 等於每天擦掉考卷,無法對答案。
  這支讓它長成時間序列:每天每檔一行(日期+訊號+收盤+PE),才能事後回測對錯。

冪等:同一天重跑會先移除當日舊列再寫,不重複累計。
資料源:data/PE買入區間監看.xlsx 的「全部監看」分頁(已含 訊號/收盤/PER/PE位階/評等)。
輸出  :data/daily_signal_log.csv(append-only 時間序列)
"""
import os
from datetime import datetime
import pandas as pd

SRC = "data/PE買入區間監看.xlsx"
LOG = "data/daily_signal_log.csv"
SHEET = "全部監看"

# 記錄欄位:訊號是「推薦結論」,收盤/PER 是「當下狀態」,事後接報酬就能驗證
COLS = {
    "訊號": "訊號",
    "未來訊號": "未來訊號",
    "評等": "評等",
    "收盤": "收盤",
    "PER現(自算)": "PER",
    "PE位階%": "PE位階",
    "PBR位階": "PBR位階",
    "含金量": "含金量",
}


def main():
    if not os.path.exists(SRC):
        print(f"⚠️ 找不到 {SRC},略過(PE 監看可能尚未產出)")
        return
    df = pd.read_excel(SRC, SHEET)
    df["代號"] = df["代號"].astype(str)
    today = datetime.now().strftime("%Y-%m-%d")

    rec = pd.DataFrame({"日期": today, "代號": df["代號"], "名稱": df["名稱"]})
    for src_col, out_col in COLS.items():
        rec[out_col] = df[src_col] if src_col in df.columns else None

    if os.path.exists(LOG):
        old = pd.read_csv(LOG, dtype={"代號": str})
        old = old[old["日期"] != today]          # 冪等:清掉當日舊列
        rec = pd.concat([old, rec], ignore_index=True)

    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    rec.to_csv(LOG, index=False, encoding="utf-8-sig")
    ndays = rec["日期"].nunique()
    print(f"已記錄 {today}:{len(df)} 檔 → {LOG}(累計 {len(rec)} 列 / {ndays} 個交易日)")


if __name__ == "__main__":
    main()
