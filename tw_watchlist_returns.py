# -*- coding: utf-8 -*-
"""
台股 A 級 120 檔 - 10/5/3/1y 含息年化批量計算 tw_watchlist_returns.py
=======================================================================
從 data/台股_體檢總表.xlsx 的「A級好公司」分頁讀 120 檔,
用 FinMind TaiwanStockPriceAdj(已含股息再投資)算總報酬 + 年化 CAGR。
輸出 data/台股_A級長期報酬榜.xlsx,按 10y 年化排序。

跑法:  python tw_watchlist_returns.py
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
SRC = "data/台股_體檢總表.xlsx"
DST = "data/台股_A級長期報酬榜.xlsx"
WORKERS = int(os.environ.get("RET_WORKERS", "4"))


def get(dataset, **params):
    params["dataset"] = dataset
    if TOKEN: params["token"] = TOKEN
    for attempt in range(3):
        try:
            r = requests.get(BASE, params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 402: time.sleep(5); continue
        if r.status_code != 200: return None
        j = r.json()
        if j.get("status") != 200: return None
        return j.get("data", [])
    return None


def cagr_for(code):
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=int(11*365.25))).isoformat()
    data = get("TaiwanStockPriceAdj", data_id=code, start_date=start)
    if not data: data = get("TaiwanStockPrice", data_id=code, start_date=start)
    if not data: return code, None
    df = pd.DataFrame(data).sort_values("date").reset_index(drop=True)
    if "close" not in df.columns or len(df) == 0: return code, None
    df["date"] = pd.to_datetime(df["date"])
    df["adjClose"] = df["close"]

    last_d, last_p = df.iloc[-1]["date"], df.iloc[-1]["adjClose"]
    out = {"代號": code, "現價(調)": round(last_p, 2)}
    for label, yrs in [("10y", 10), ("5y", 5), ("3y", 3), ("1y", 1)]:
        tgt = last_d - pd.Timedelta(days=int(yrs*365.25))
        sub = df[df["date"] <= tgt]
        if len(sub) == 0 or sub.iloc[-1]["adjClose"] <= 0:
            out[f"{label}年化%"] = None; out[f"{label}總報酬%"] = None; continue
        p0 = sub.iloc[-1]["adjClose"]
        total = (last_p / p0 - 1) * 100
        cagr = ((last_p / p0) ** (1/yrs) - 1) * 100
        out[f"{label}總報酬%"] = round(total, 1)
        out[f"{label}年化%"] = round(cagr, 2)
    return code, out


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN(會被速率限制)")
    # 讀 A 級清單
    sheets = pd.read_excel(SRC, sheet_name=None)
    a = sheets.get("A級好公司")
    if a is None or len(a) == 0:
        print("⚠️ 找不到 A級好公司 分頁,改用主表評等=A")
        a = sheets["體檢總表"]
        a = a[a["評等"] == "A"]
    a["代號"] = a["代號"].astype(str)
    codes = a["代號"].tolist()
    print(f"算 {len(codes)} 檔 10/5/3/1y 含息年化(平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(cagr_for, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            code, data = fut.result()
            if data: results[code] = data
            done += 1
            if done % 20 == 0: print(f"  [{done}/{len(codes)}]")

    df = pd.DataFrame(list(results.values()))
    keep = ["代號","名稱","產業","評等","品質總分"]
    keep = [c for c in keep if c in a.columns]
    df = df.merge(a[keep], on="代號", how="left")

    df = df.sort_values("10y年化%", ascending=False, na_position="last")
    cols = ["代號","名稱","產業","評等","品質總分","現價(調)",
            "10y年化%","10y總報酬%","5y年化%","5y總報酬%",
            "3y年化%","3y總報酬%","1y年化%","1y總報酬%"]
    df = df[[c for c in cols if c in df.columns]]

    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="長期報酬榜", index=False)
        df.head(20).to_excel(xw, sheet_name="TOP20_10y年化", index=False)
        df.sort_values("1y年化%").head(20).to_excel(xw, sheet_name="近1y跌幅榜", index=False)

    print(f"\n→ 已輸出 {DST}\n")
    print("=== 10y 含息年化 TOP 15 ===")
    show = ["代號","名稱","品質總分","10y年化%","5y年化%","3y年化%","1y年化%"]
    show = [c for c in show if c in df.columns]
    print(df[show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
