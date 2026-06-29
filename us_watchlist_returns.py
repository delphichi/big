# -*- coding: utf-8 -*-
"""
監看清單 96 檔 - 10/5/3/1y 含息年化報酬批量計算 us_watchlist_returns.py
=======================================================================
用 FMP adjClose(已含股息再投資)算總報酬 + 年化 CAGR
輸出 data/美股長期報酬榜.xlsx,按 10y 年化排序

跑法:  python us_watchlist_returns.py
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
SRC = "data/美股體檢總表.xlsx"
DST = "data/美股長期報酬榜.xlsx"
WORKERS = int(os.environ.get("RET_WORKERS", "6"))

DEFAULT_WATCH = """
NVDA NVMI ANET KLAC AVGO FTNT ASML TSM LLY BRK-B MSFT META NFLX
WPM AXP AMZN CDNS CAT HWM AAPL APH GLW GOOG EXEL
GLD HG CCEP CF AER LIN AMG COST ECL WMT FSLR IDCC NBIX
CLS LRCX AMD MU MRVL LITE AMAT CIEN COHR WDC VRT PLTR
CRM ADBE INTU NOW
GE MCO SPGI CP FER CNI
""".split()


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1)); continue
        if r.status_code != 200:
            return None
        try: return r.json()
        except: return None
    return None


def cagr_for(sym):
    today = datetime.now(timezone.utc).date()
    start = today - timedelta(days=int(11*365.25))
    data = get("historical-price-eod/full", symbol=sym,
               **{"from": start.isoformat(), "to": today.isoformat()})
    if not data: return sym, None
    hist = data.get("historical") if isinstance(data, dict) else data
    if not hist: return sym, None
    df = pd.DataFrame(hist)
    if "adjClose" not in df.columns: df["adjClose"] = df.get("close")
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
    if len(df) == 0: return sym, None

    last_d = df.iloc[-1]["date"]
    last_p = df.iloc[-1]["adjClose"]
    out = {"代號": sym, "現價(調)": round(last_p, 2)}
    for label, yrs in [("10y", 10), ("5y", 5), ("3y", 3), ("1y", 1)]:
        tgt = last_d - pd.Timedelta(days=int(yrs*365.25))
        sub = df[df["date"] <= tgt]
        if len(sub) == 0:
            out[f"{label}年化%"] = None; out[f"{label}總報酬%"] = None; continue
        p0 = sub.iloc[-1]["adjClose"]
        if p0 <= 0:
            out[f"{label}年化%"] = None; out[f"{label}總報酬%"] = None; continue
        total = (last_p / p0 - 1) * 100
        cagr = ((last_p / p0) ** (1/yrs) - 1) * 100
        out[f"{label}總報酬%"] = round(total, 1)
        out[f"{label}年化%"] = round(cagr, 2)
    return sym, out


def main():
    if not KEY: print("⚠️ 未設 FMP_API_KEY"); return
    syms = DEFAULT_WATCH
    print(f"算 {len(syms)} 檔 10/5/3/1y 含息年化(平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(cagr_for, s): s for s in syms}
        done = 0
        for fut in as_completed(futs):
            sym, data = fut.result()
            if data: results[sym] = data
            done += 1
            if done % 20 == 0: print(f"  [{done}/{len(syms)}]")

    rows = list(results.values())
    df = pd.DataFrame(rows)

    # merge 名稱 / 評等 / 品質
    base = pd.read_excel(SRC, sheet_name="體檢總表")[["代號","名稱","產業","評等","品質總分"]]
    base["代號"] = base["代號"].astype(str)
    df = df.merge(base, on="代號", how="left")

    # 排序按 10y 年化
    df = df.sort_values("10y年化%", ascending=False, na_position="last")

    cols = ["代號","名稱","產業","評等","品質總分","現價(調)",
            "10y年化%","10y總報酬%","5y年化%","5y總報酬%",
            "3y年化%","3y總報酬%","1y年化%","1y總報酬%"]
    df = df[[c for c in cols if c in df.columns]]

    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="長期報酬榜", index=False)
        df.head(20).to_excel(xw, sheet_name="TOP20_10y年化", index=False)
        # 1y 跌幅榜(可能是低接機會)
        bot = df.sort_values("1y年化%").head(20)
        bot.to_excel(xw, sheet_name="近1y跌幅榜", index=False)

    print(f"\n→ 已輸出 {DST}\n")
    print("=== 10y 含息年化 TOP 15 ===")
    print(df[["代號","名稱","評等","10y年化%","5y年化%","3y年化%","1y年化%"]].head(15).to_string(index=False))
    print("\n=== 10y 含息年化 BOTTOM 10 ===")
    print(df[["代號","名稱","評等","10y年化%","5y年化%","3y年化%","1y年化%"]].tail(10).to_string(index=False))


if __name__ == "__main__":
    main()
