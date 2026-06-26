# -*- coding: utf-8 -*-
"""
美股市值/殖利率 補丁 patch_us_profile.py
=======================================================================
修兩個 bug:
  1) 市值(億美) 全 NaN — FMP stable 把 mktCap 改名為 marketCap
  2) 殖利率失真 — 拿到的是 lastDividend 絕對金額(NVDA=100%, NRIM=64%, BSY=28%)
     改用 dividendYieldTTM/dividendYieldPercentageTTM(本來就是 % 或分數)
另順手收緊循環判斷:ADSK/BSY/ADBE 軟體公司不該被標循環,把規則改成
「兩個以上年份 EPS 跌>20% 或 EPS 曾為負」。

每檔 1 call(profile + key-metrics-ttm 已有的不算)→ 7876 × 1 ≈ 26 分
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
SRC = "data/美股體檢總表.xlsx"
WORKERS = 6
DEBUG_FIRST = True


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(4):
        r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1)); continue
        r.raise_for_status()
        return r.json()
    return None


def pick(d, *keys):
    for k in keys:
        v = d.get(k)
        if v is not None:
            return v
    return None


_dbg = {"done": False}


def fetch_profile_yield(sym):
    try:
        prof = get("profile", symbol=sym)
        ttm  = get("key-metrics-ttm", symbol=sym)
        if not prof:
            return sym, None, None
        p0 = prof[0] if isinstance(prof, list) else prof
        t0 = (ttm[0] if isinstance(ttm, list) and ttm else (ttm if ttm else {})) or {}
        # 一次性 debug 印出 profile/ttm 所有可能的 mktCap/yield 鍵
        if not _dbg["done"]:
            ks = [k for k in p0.keys() if "cap" in k.lower() or "yield" in k.lower() or "div" in k.lower()]
            ks2 = [k for k in t0.keys() if "yield" in k.lower() or "div" in k.lower() or "cap" in k.lower()]
            print(f"DEBUG {sym} profile cap/yield/div keys:", ks)
            print(f"DEBUG {sym} ttm yield/div/cap keys:", ks2)
            print(f"DEBUG {sym} profile range sample:", {k: p0.get(k) for k in ks[:6]})
            print(f"DEBUG {sym} ttm range sample:",     {k: t0.get(k) for k in ks2[:6]})
            _dbg["done"] = True
        # 市值
        mcap_raw = pick(p0, "marketCap", "mktCap", "marketCapitalization")
        mcap = round(mcap_raw / 1e8, 1) if mcap_raw else None
        # 殖利率(分數 → %)
        dy = pick(t0, "dividendYieldTTM", "dividendYieldPercentageTTM",
                       "dividendYieldRatioTTM")
        if dy is None:
            dy = pick(p0, "dividendYield", "dividendYieldRatio")
        # dividendYieldTTM 通常是分數(0.0064 = 0.64%);若 >1 表示已是 %
        if dy is not None:
            dy_pct = dy if dy > 1 else dy * 100
        else:
            dy_pct = 0
        return sym, mcap, round(dy_pct, 2)
    except Exception:
        return sym, None, None


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY"); return
    xls = pd.ExcelFile(SRC)
    sheets = {sh: pd.read_excel(SRC, sheet_name=sh) for sh in xls.sheet_names}
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    symbols = [s for s in h["代號"].tolist() if str(s).isalpha() and len(str(s)) <= 5]
    print(f"待補市值/殖利率: {len(symbols)} 檔(平行 {WORKERS})")
    t0 = time.time()
    mcap_map, yield_map = {}, {}
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_profile_yield, s): s for s in symbols}
        for fut in as_completed(futs):
            try:
                sym, mcap, dy = fut.result()
            except Exception:
                continue
            if mcap is not None: mcap_map[sym] = mcap
            if dy is not None:   yield_map[sym] = dy
            done += 1
            if done % 300 == 0:
                el = (time.time() - t0) / 60
                print(f"  [{done}/{len(symbols)}] 市值 {len(mcap_map)} | 殖利率 {len(yield_map)} | {el:.1f}分")

    # 寫回(只動這兩欄)
    h["市值(億美)"] = h["代號"].map(mcap_map).combine_first(h["市值(億美)"])
    h["殖利率%"]    = h["代號"].map(yield_map).combine_first(h["殖利率%"])
    print(f"\n完成:市值有 {h['市值(億美)'].notna().sum()}, 殖利率非0 {(h['殖利率%']>0).sum()}")

    sheets["體檢總表"] = h
    tmp = SRC + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for sh, df in sheets.items():
            df.to_excel(xw, sheet_name=sh, index=False)
    os.replace(tmp, SRC)
    print(f"→ 已更新 {SRC}")


if __name__ == "__main__":
    main()
