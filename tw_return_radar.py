# -*- coding: utf-8 -*-
"""
台股報酬雷達 tw_return_radar.py
=======================================================================
對 watchlist 用 TaiwanStockPriceAdj (還原股價) 算:
  - 1m / 3m / 6m / 1y / 3y / 5y / 10y 報酬
  - 對應期間「年化報酬」
  - 對標大盤 (TaiwanStockTotalReturnIndex TAIEX) 算「超額報酬」
  - 最大回撤 (max drawdown)

Watchlist 來源: TICKERS env → data/watchlist_tw.txt → fallback

輸出 data/台股_報酬雷達.xlsx, 3 個分頁:
  - 總覽 (每檔 7 期報酬 + 超額)
  - 超額排行
  - 大盤對照 (TAIEX 報酬指數)
"""
import os, time, requests
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股_報酬雷達.xlsx"
WATCHLIST_FILE = "data/watchlist_tw.txt"
WORKERS = int(os.environ.get("WORKERS", "4"))
END = datetime.now().strftime("%Y-%m-%d")
START = (datetime.now() - timedelta(days=365*11)).strftime("%Y-%m-%d")  # 11 年 buffer


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
        toks = [t for t in toks if t and not t.startswith("#")]
        if toks: return list(dict.fromkeys(toks))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip() for t in line.split() if t.strip())
        if toks: return list(dict.fromkeys(toks))
    return "2330 2454 2317".split()


def fm(dataset, data_id=None, start=START, end=END):
    p = {"dataset": dataset, "start_date": start, "end_date": end}
    if data_id: p["data_id"] = data_id
    if TOKEN: p["token"] = TOKEN
    for _ in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=30)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception:
            time.sleep(1)
    return pd.DataFrame()


# ─── 全市場 一次抓 TAIEX 報酬指數做 benchmark ───
TAIEX_PRICES = None
def get_taiex():
    global TAIEX_PRICES
    if TAIEX_PRICES is None:
        df = fm("TaiwanStockTotalReturnIndex", data_id="TAIEX")
        if df.empty: TAIEX_PRICES = pd.DataFrame()
        else:
            df = df.sort_values("date").reset_index(drop=True)
            df["date"] = pd.to_datetime(df["date"])
            TAIEX_PRICES = df.set_index("date")["price"]
    return TAIEX_PRICES


def return_at(prices, days):
    """從最新一筆往前回推 days 天的收盤, 算累計報酬%"""
    if len(prices) < 2: return None
    end_p = prices.iloc[-1]
    target_date = prices.index[-1] - timedelta(days=days)
    # 取離 target_date 最近的有效資料
    past = prices[prices.index <= target_date]
    if len(past) == 0: return None
    start_p = past.iloc[-1]
    if start_p <= 0: return None
    return round((end_p / start_p - 1) * 100, 1)


def annualized(total_pct, years):
    """總報酬% → 年化%"""
    if total_pct is None or years <= 0: return None
    try: return round(((1 + total_pct/100) ** (1/years) - 1) * 100, 1)
    except: return None


def max_drawdown(prices):
    """最大回撤%"""
    if len(prices) < 2: return None
    roll_max = prices.cummax()
    dd = (prices - roll_max) / roll_max * 100
    return round(dd.min(), 1)


def fetch_one(sid):
    try:
        df = fm("TaiwanStockPriceAdj", data_id=sid)
        if df.empty or "close" not in df.columns:
            return sid, {"代號": sid, "__error": "no data"}
        df = df.sort_values("date")
        df["date"] = pd.to_datetime(df["date"])
        prices = df.set_index("date")["close"].astype(float)
        prices = prices[prices > 0]
        if len(prices) < 30: return sid, {"代號": sid, "__error": "資料不足"}

        out = {"代號": sid, "最新價": round(prices.iloc[-1], 2),
               "資料起": str(prices.index[0].date()),
               "資料迄": str(prices.index[-1].date())}

        # 7 期累計報酬
        periods = [("1m", 30, 1/12), ("3m", 90, 1/4), ("6m", 180, 1/2),
                   ("1y", 365, 1), ("3y", 365*3, 3), ("5y", 365*5, 5), ("10y", 365*10, 10)]
        for name, days, years in periods:
            r = return_at(prices, days)
            out[f"{name}報酬%"] = r
            if years >= 1:
                out[f"{name}年化%"] = annualized(r, years)
            # 對應 TAIEX
            tx = get_taiex()
            if not tx.empty:
                tr = return_at(tx, days)
                if r is not None and tr is not None:
                    out[f"{name}超額%"] = round(r - tr, 1)

        # 最大回撤 (5 年)
        cutoff = prices.index[-1] - timedelta(days=365*5)
        recent5 = prices[prices.index >= cutoff]
        out["5y最大回撤%"] = max_drawdown(recent5) if len(recent5) >= 2 else None

        return sid, out
    except Exception as e:
        return sid, {"代號": sid, "__error": str(e)}


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN")
    codes = load_watchlist()
    print(f"台股報酬雷達 — {len(codes)} 檔 (平行 {WORKERS})")

    # 預載 TAIEX
    print("預載 TAIEX 大盤報酬指數...")
    tx = get_taiex()
    if not tx.empty:
        # 大盤同期報酬
        bench = {"信號": "TAIEX (大盤總報酬)"}
        for name, days, _ in [("1m",30,0),("3m",90,0),("6m",180,0),("1y",365,0),
                              ("3y",365*3,0),("5y",365*5,0),("10y",365*10,0)]:
            bench[f"{name}%"] = return_at(tx, days)
        print(f"  TAIEX 報酬: {bench}")

    # 個股 fetch
    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sid, data = fut.result()
            if data: results[sid] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # merge 體檢
    base = pd.DataFrame()
    for src in ["data/台股體檢總表.xlsx", "data/台股_體檢總表.xlsx"]:
        if not os.path.exists(src): continue
        try:
            h = pd.read_excel(src, sheet_name="體檢總表")
            h["代號"] = h["代號"].astype(str)
            base = h[[c for c in ["代號","名稱","產業","評等","品質總分"] if c in h.columns]]
            break
        except Exception: pass

    rows = [r for r in results.values() if r and "__error" not in r]
    df = pd.DataFrame(rows)
    if not base.empty and not df.empty:
        df["代號"] = df["代號"].astype(str)
        df = df.merge(base, on="代號", how="left")
        front = [c for c in ["代號","名稱","產業","評等","品質總分","最新價",
                              "1m報酬%","3m報酬%","6m報酬%","1y報酬%","3y報酬%","5y報酬%","10y報酬%",
                              "1y年化%","3y年化%","5y年化%","10y年化%",
                              "1m超額%","3m超額%","6m超額%","1y超額%","3y超額%","5y超額%","10y超額%",
                              "5y最大回撤%"] if c in df.columns]
        rest = [c for c in df.columns if c not in front]
        df = df[front + rest]
    df = df.sort_values("1y報酬%", ascending=False, na_position="last")

    # 超額排行(看 1y/3y/5y 超額)
    excess_cols = [c for c in ["代號","名稱","評等","1y超額%","3y超額%","5y超額%","10y超額%"] if c in df.columns]
    excess_sheet = df[excess_cols].copy().sort_values("3y超額%", ascending=False, na_position="last")

    # TAIEX benchmark sheet
    tx_df = pd.DataFrame()
    if not tx.empty:
        rows = []
        for name, days, years in [("1m",30,1/12),("3m",90,1/4),("6m",180,1/2),
                                   ("1y",365,1),("3y",365*3,3),("5y",365*5,5),("10y",365*10,10)]:
            r = return_at(tx, days)
            rows.append({"期間": name, "累計報酬%": r, "年化%": annualized(r, years) if years>=1 else None})
        tx_df = pd.DataFrame(rows)

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="總覽", index=False)
        excess_sheet.to_excel(xw, sheet_name="超額排行", index=False)
        if not tx_df.empty:
            tx_df.to_excel(xw, sheet_name="TAIEX大盤", index=False)

    print(f"\n→ {DST}")
    print(f"\n=== 1y 報酬 TOP 15 ===")
    show = [c for c in ["代號","名稱","評等","1y報酬%","3y年化%","5y年化%","3y超額%","5y超額%","5y最大回撤%"] if c in df.columns]
    print(df[show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
