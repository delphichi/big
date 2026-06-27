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
5511 1514 2645 2472 2535 8926 2207 3596 5225 5243 1616 2630 3171 2476 2451
4933 6274 3303 3188 3260 2607 3030 2317 4904 6690 2605 6506 3402 4569 3147
6739 3162 2504 6509 1795 2305 3492 4560 4303 6263 6435 3628 2910 1785 1608
3227 9911 2752 3652 2312 5206 3128 2471 1590 3025 4105 6526 7556 4175 1730
2480 2754 3209 4728 9942 2528 1773 6419 3081 2612 1476 9917 2727 4155 6166
6525 4754 6192 4951 1777 8271 5902 6640 2393 3406 4207 6741 3711 2488 2423
2453 6257 1580 5381 1110 2356 2707 2751 1712 2731 4987 6482 2762 1615 4744
2101 3218 6446 5522 1726 6510 4916 4931 4979 6532 3293 3605 4546 2382 1503
1708 3498 3705 5287 6944 4581 3633 5209 3622 2301 3019 3687 6712 2316 2412
3661 3450 6291 2745 1735 1472 4541 3684 1258 2420 2006 5706 2606 2633 3416
6285 4772 8462 4771 2008 1732 2493 2901 5443 5289 3231 3673 8299 3006 4106
1720 6206 2031 6826 3265 1419 2408 5410 1470 6556 1737 3617 1524 4536 1210
5345 4991 4565 4131 4973 1752 5548 2245 2395 4749 4205 5288 2379 6498 1477
4413 2706 4720 5353 3332 6103 2327 6752 6117 6104 5529 2010 5607 1240 5269
2617 6214 1233 4114 3078 6613 6282 8086 3709 2916 1535 6231 3130 4950
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
