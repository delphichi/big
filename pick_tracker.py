# -*- coding: utf-8 -*-
"""
推薦買入追蹤 + 戰績評估 pick_tracker.py
=======================================================================
功能:
  1) record:  從台股/美股 PE 監看表撈當天「買進信號」(🟢),append 到 data/推薦追蹤.csv
  2) evaluate: 對每筆紀錄,若 1m/3m/6m/1y 時間到,自動補價格與報酬率
  3) report:   產 data/推薦戰績.xlsx,含總勝率/平均報酬/各市場分析

跑法:
  python pick_tracker.py              # 預設 record + evaluate + report
  python pick_tracker.py record       # 只 append 新推薦
  python pick_tracker.py evaluate     # 只補歷史推薦的後續價
"""
import os
import sys
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

TW_PE = "data/台股PE監看表.xlsx"
US_PE = "data/美股PE監看表.xlsx"
TRACK_CSV = "data/推薦追蹤.csv"
REPORT_XLSX = "data/推薦戰績.xlsx"

FMP_KEY = os.environ.get("FMP_API_KEY", "")
FM_TOKEN = os.environ.get("FINMIND_TOKEN", "")
FMP_BASE = "https://financialmodelingprep.com/stable"
FM_BASE = "https://api.finmindtrade.com/api/v4/data"

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).date()
BUY_TAGS = ["🟢未來便宜", "🟢成長未反映", "🟢便宜"]


# ---------- 抓即時報價 ----------
def fetch_tw_price(code):
    if not FM_TOKEN: return None
    start = (TODAY - timedelta(days=10)).isoformat()
    try:
        r = requests.get(FM_BASE, params={"dataset":"TaiwanStockPrice","data_id":code,
                                          "start_date":start,"token":FM_TOKEN}, timeout=15)
        if r.status_code != 200: return None
        j = r.json()
        if j.get("status") != 200: return None
        data = j.get("data", [])
        if not data: return None
        return float(pd.DataFrame(data).sort_values("date").iloc[-1]["close"])
    except Exception: return None


def fetch_us_price(sym):
    if not FMP_KEY: return None
    try:
        r = requests.get(f"{FMP_BASE}/quote", params={"symbol":sym,"apikey":FMP_KEY}, timeout=15)
        if r.status_code != 200: return None
        d = r.json()
        if not d: return None
        return float(d[0].get("price", 0)) or None
    except Exception: return None


def fetch_price(market, code):
    return fetch_tw_price(code) if market == "TW" else fetch_us_price(code)


# ---------- record:撈當天買進信號 ----------
def collect_picks():
    picks = []
    for src, market in [(TW_PE, "TW"), (US_PE, "US")]:
        if not os.path.exists(src): continue
        try:
            df = pd.read_excel(src, sheet_name="監看表")
        except Exception:
            continue
        df["代號"] = df["代號"].astype(str)
        buy = df[df["估值鬧鐘"].isin(BUY_TAGS)].copy()
        if len(buy) == 0: continue
        price_col = "當前股價"
        for _, r in buy.iterrows():
            price = r.get(price_col)
            if pd.isna(price) or not price: continue
            picks.append({
                "推薦日": TODAY.isoformat(),
                "市場": market,
                "代號": str(r["代號"]),
                "名稱": r.get("名稱", ""),
                "推薦價": round(float(price), 2),
                "推薦類型": r.get("估值鬧鐘", ""),
                "評等": r.get("評等", ""),
                "品質總分": r.get("品質總分"),
                "ForwardPE": r.get("ForwardPE即時") if "ForwardPE即時" in r else r.get("ForwardPE"),
                "PEG": r.get("PEG即時") if "PEG即時" in r else r.get("PEG_使用"),
                "1m_價": None, "1m_報酬%": None,
                "3m_價": None, "3m_報酬%": None,
                "6m_價": None, "6m_報酬%": None,
                "1y_價": None, "1y_報酬%": None,
            })
    return picks


def record_picks():
    new = collect_picks()
    if not new:
        print("今天無買進信號"); return
    new_df = pd.DataFrame(new)
    # 去重:同代號 7 天內已記錄過就跳過
    if os.path.exists(TRACK_CSV):
        old = pd.read_csv(TRACK_CSV, dtype={"代號": str})
        old["推薦日"] = pd.to_datetime(old["推薦日"]).dt.date
        cutoff = TODAY - timedelta(days=7)
        recent = old[old["推薦日"] >= cutoff][["市場","代號"]].apply(tuple, axis=1).tolist()
        new_df = new_df[~new_df.apply(lambda r: (r["市場"], r["代號"]) in recent, axis=1)]
    if len(new_df) == 0:
        print("所有推薦 7 天內已記錄過,跳過"); return
    header = not os.path.exists(TRACK_CSV)
    new_df.to_csv(TRACK_CSV, mode="a", header=header, index=False)
    print(f"→ 新增 {len(new_df)} 筆推薦到 {TRACK_CSV}")
    print(new_df[["市場","代號","名稱","推薦價","推薦類型"]].to_string(index=False))


# ---------- evaluate:補時間到的後續價格 ----------
def evaluate_picks():
    if not os.path.exists(TRACK_CSV):
        print("還沒任何推薦紀錄"); return
    df = pd.read_csv(TRACK_CSV, dtype={"代號": str})
    df["推薦日"] = pd.to_datetime(df["推薦日"]).dt.date

    horizons = [("1m", 30), ("3m", 90), ("6m", 180), ("1y", 365)]
    updates = 0

    for idx, row in df.iterrows():
        for label, days in horizons:
            price_col = f"{label}_價"
            ret_col = f"{label}_報酬%"
            if pd.notna(row[price_col]): continue  # 已填過
            target_date = row["推薦日"] + timedelta(days=days)
            if TODAY < target_date: continue  # 還沒到
            price = fetch_price(row["市場"], row["代號"])
            if not price: continue
            df.at[idx, price_col] = round(price, 2)
            ret = (price / row["推薦價"] - 1) * 100
            df.at[idx, ret_col] = round(ret, 2)
            updates += 1
    if updates:
        df.to_csv(TRACK_CSV, index=False)
        print(f"→ 補了 {updates} 個價格點")
    else:
        print("沒有時間到的推薦要補")


# ---------- report:產戰績摘要 ----------
def report():
    if not os.path.exists(TRACK_CSV):
        print("沒有推薦紀錄"); return
    df = pd.read_csv(TRACK_CSV, dtype={"代號": str})
    df["推薦日"] = pd.to_datetime(df["推薦日"]).dt.date

    summary = {"區間": [], "已評估筆數": [], "勝率%": [],
               "平均報酬%": [], "中位數%": [], "最佳%": [], "最差%": []}
    for label in ["1m", "3m", "6m", "1y"]:
        col = f"{label}_報酬%"
        vals = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(vals) == 0:
            summary["區間"].append(label); summary["已評估筆數"].append(0)
            for k in ["勝率%","平均報酬%","中位數%","最佳%","最差%"]: summary[k].append(None)
            continue
        wins = (vals > 0).sum()
        summary["區間"].append(label)
        summary["已評估筆數"].append(len(vals))
        summary["勝率%"].append(round(wins / len(vals) * 100, 1))
        summary["平均報酬%"].append(round(vals.mean(), 2))
        summary["中位數%"].append(round(vals.median(), 2))
        summary["最佳%"].append(round(vals.max(), 2))
        summary["最差%"].append(round(vals.min(), 2))

    sm = pd.DataFrame(summary)

    # 個別市場 / 推薦類型分析
    by_market = []
    for m in df["市場"].unique():
        sub = df[df["市場"] == m]
        for label in ["1m","3m","6m","1y"]:
            vals = pd.to_numeric(sub[f"{label}_報酬%"], errors="coerce").dropna()
            if len(vals)==0: continue
            by_market.append({"市場": m, "區間": label, "筆數": len(vals),
                              "勝率%": round((vals>0).sum()/len(vals)*100, 1),
                              "平均報酬%": round(vals.mean(), 2)})
    bm = pd.DataFrame(by_market)

    # 排行榜
    df_eval = df.copy()
    df_eval["最終報酬%"] = df_eval[["1y_報酬%","6m_報酬%","3m_報酬%","1m_報酬%"]].bfill(axis=1).iloc[:,0]
    top = df_eval[df_eval["最終報酬%"].notna()].sort_values("最終報酬%", ascending=False).head(20)
    bot = df_eval[df_eval["最終報酬%"].notna()].sort_values("最終報酬%").head(20)

    with pd.ExcelWriter(REPORT_XLSX, engine="openpyxl") as xw:
        sm.to_excel(xw, sheet_name="總戰績", index=False)
        bm.to_excel(xw, sheet_name="分市場戰績", index=False)
        df.to_excel(xw, sheet_name="完整紀錄", index=False)
        top.to_excel(xw, sheet_name="最佳20", index=False)
        bot.to_excel(xw, sheet_name="最差20", index=False)

    print(f"\n→ 已輸出 {REPORT_XLSX}\n")
    print("=== 總戰績 ==="); print(sm.to_string(index=False))
    if len(bm): print("\n=== 分市場 ==="); print(bm.to_string(index=False))
    print(f"\n累積推薦 {len(df)} 筆,已有 1m 評估 {sm['已評估筆數'].iloc[0]} 筆")


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("record", "all"):  record_picks()
    if mode in ("evaluate", "all"): evaluate_picks()
    if mode in ("report", "all"):   report()


if __name__ == "__main__":
    main()
