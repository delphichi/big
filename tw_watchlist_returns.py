# -*- coding: utf-8 -*-
"""
台股自選 Watchlist - 10/5/3/1y 含息年化批量計算 tw_watchlist_returns.py
=======================================================================
用 FinMind TaiwanStockPriceAdj(已含股息再投資)算總報酬 + 年化 CAGR。
輸出 data/台股_長期報酬榜.xlsx,按 10y 年化排序。

跑法:  python tw_watchlist_returns.py
切換清單:
  WATCHLIST=2330,3017,6139 python tw_watchlist_returns.py
  (空白則用內建 DEFAULT_WATCH)
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
DST = "data/台股_長期報酬榜.xlsx"
WORKERS = int(os.environ.get("RET_WORKERS", "4"))

DEFAULT_WATCH = """
3017 3653 2383 6139 8210 8996 6442 2360 2345 6223 3324
2330 2308 6196 3583 6197 3044 2059 5434 1513 1560 5340 1519 3008
2640 3029 4506 3689 5519 2421 1618 1215 2618 5904 1232 6189 5478 4527
6788 2753 6515 3045 4129 2912 1514 3147 3402
6274 2472 3260 3406 3661 2327 2408 3711 3293 1503 2535 2476 6263 1773 1590
4303 5269 6446 4979 3450 3081 8271 2382 2379 3231 2301
5511 8926 1616 4933 3596 3227 6285 6510 2356 2317 5243 5225 3188 2305 8086
2395 2412 9917 6739 2645 6690 4569 4904 2344
""".split()


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

    watch = os.environ.get("WATCHLIST", "").strip()
    if watch:
        codes = [s.strip() for s in watch.replace(",", " ").split() if s.strip()]
        print(f"(自訂清單)")
    else:
        codes = DEFAULT_WATCH
    codes = list(dict.fromkeys(codes))  # 去重保序
    print(f"算 {len(codes)} 檔 10/5/3/1y 含息年化(平行 {WORKERS})")

    # 從體檢總表撈 名稱/產業/評等/品質總分
    try:
        base = pd.read_excel(SRC, sheet_name="體檢總表")
        base["代號"] = base["代號"].astype(str)
        keep = [c for c in ["代號","名稱","產業","評等","品質總分"] if c in base.columns]
        base = base[keep]
    except Exception as e:
        print(f"⚠️ 讀體檢總表失敗 {e},只回價格資料")
        base = pd.DataFrame({"代號": codes})

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(cagr_for, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            code, data = fut.result()
            if data: results[code] = data
            done += 1
            if done % 30 == 0: print(f"  [{done}/{len(codes)}]")

    df = pd.DataFrame(list(results.values()))
    df = df.merge(base, on="代號", how="left")

    df = df.sort_values("10y年化%", ascending=False, na_position="last")
    cols = ["代號","名稱","產業","評等","品質總分","現價(調)",
            "10y年化%","10y總報酬%","5y年化%","5y總報酬%",
            "3y年化%","3y總報酬%","1y年化%","1y總報酬%"]
    df = df[[c for c in cols if c in df.columns]]

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="長期報酬榜", index=False)
        df.head(30).to_excel(xw, sheet_name="TOP30_10y年化", index=False)
        df.sort_values("1y年化%").head(30).to_excel(xw, sheet_name="近1y跌幅榜", index=False)
        # 跌深A級(長期強+近期跌)
        if "評等" in df.columns:
            ab = df[df["評等"].isin(["A","B"])].copy()
            mask = (ab["10y年化%"] >= 10) & (ab["1y年化%"] <= 0)
            opp = ab[mask].sort_values("1y年化%")
            if len(opp): opp.to_excel(xw, sheet_name="跌深A_B級低接候選", index=False)

    print(f"\n→ 已輸出 {DST}\n")
    print("=== 10y 含息年化 TOP 20 ===")
    show = [c for c in ["代號","名稱","評等","品質總分","10y年化%","5y年化%","3y年化%","1y年化%"] if c in df.columns]
    print(df[show].head(20).to_string(index=False))
    print(f"\n=== 近 1y 跌幅 TOP 15 ===")
    print(df.sort_values("1y年化%")[show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
