# -*- coding: utf-8 -*-
"""
美股市值/殖利率 補丁 patch_us_profile.py
=======================================================================
修兩個 bug:
  1) 市值(億美) 全 NaN — FMP stable 把 mktCap 改名為 marketCap
  2) 殖利率失真(>1 表示被當成絕對金額)

預設只 patch A+B 級(639 檔, 5 分內完成);要全 patch 設 ALL=1
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
SRC = "data/美股體檢總表.xlsx"
WORKERS = int(os.environ.get("PATCH_WORKERS", "4"))
MAX_RUNTIME_MIN = int(os.environ.get("PATCH_MAX_MIN", "38"))   # 早於 CI 45 分 timeout
ALL_GRADES = os.environ.get("ALL", "") == "1"


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=15)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1)); continue
        if r.status_code >= 500:
            time.sleep(1); continue
        if r.status_code != 200:
            return None
        try: return r.json()
        except: return None
    return None


def pick(d, *keys):
    for k in keys:
        v = d.get(k)
        if v is not None:
            return v
    return None


def fetch_profile_yield(sym):
    try:
        prof = get("profile", symbol=sym)
        if not prof: return sym, None, None
        p0 = prof[0] if isinstance(prof, list) else prof
        ttm  = get("key-metrics-ttm", symbol=sym)
        t0 = ((ttm[0] if isinstance(ttm, list) and ttm else (ttm or {}))) or {}
        # 市值
        mcap_raw = pick(p0, "marketCap", "mktCap", "marketCapitalization")
        mcap = round(mcap_raw / 1e8, 1) if mcap_raw else None
        # 殖利率
        dy = pick(t0, "dividendYieldTTM", "dividendYieldPercentageTTM", "dividendYieldRatioTTM")
        if dy is None:
            dy = pick(p0, "dividendYield", "dividendYieldRatio")
        if dy is not None:
            dy_pct = dy if dy > 1 else dy * 100
        else:
            dy_pct = 0
        return sym, mcap, round(dy_pct, 2)
    except Exception:
        return sym, None, None


def write_back(sheets, mcap_map, yield_map, src):
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    if mcap_map:
        m = h["代號"].map(mcap_map)
        h["市值(億美)"] = m.combine_first(h["市值(億美)"])
    if yield_map:
        m = h["代號"].map(yield_map)
        h["殖利率%"] = m.combine_first(h["殖利率%"])
    sheets["體檢總表"] = h
    tmp = src + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for sh, df in sheets.items():
            df.to_excel(xw, sheet_name=sh, index=False)
    os.replace(tmp, src)


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY"); return
    xls = pd.ExcelFile(SRC)
    sheets = {sh: pd.read_excel(SRC, sheet_name=sh) for sh in xls.sheet_names}
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    if ALL_GRADES:
        target = h
        print(f"模式:全部 {len(h)} 檔")
    else:
        target = h[h["評等"].isin(["A", "B"])]
        print(f"模式:只 patch A+B 級 {len(target)} 檔")
    symbols = [s for s in target["代號"].tolist() if str(s).isalpha() and len(str(s)) <= 5]
    print(f"待 patch: {len(symbols)} 檔(平行 {WORKERS}, 自停 {MAX_RUNTIME_MIN}分)")

    t0 = time.time()
    mcap_map, yield_map = {}, {}
    done = 0
    ex = ThreadPoolExecutor(max_workers=WORKERS)
    futs = {ex.submit(fetch_profile_yield, s): s for s in symbols}
    for fut in as_completed(futs):
        try:
            sym, mcap, dy = fut.result()
        except Exception:
            continue
        if mcap is not None: mcap_map[sym] = mcap
        if dy is not None:   yield_map[sym] = dy
        done += 1
        if done % 50 == 0:
            el = (time.time() - t0) / 60
            print(f"  [{done}/{len(symbols)}] 市值 {len(mcap_map)} | 殖利率 {len(yield_map)} | {el:.1f}分")
        if done % 500 == 0:
            write_back(sheets, mcap_map, yield_map, SRC)
            print("    中途存檔")
        if (time.time() - t0) / 60 > MAX_RUNTIME_MIN:
            print(f"⏲ 達 {MAX_RUNTIME_MIN} 分自停,取消未開始的並收尾")
            ex.shutdown(wait=False, cancel_futures=True)
            break
    ex.shutdown(wait=True)

    write_back(sheets, mcap_map, yield_map, SRC)
    print(f"\n完成:市值 {len(mcap_map)}, 殖利率非0 {sum(1 for v in yield_map.values() if v>0)}")
    print(f"→ 已更新 {SRC}")


if __name__ == "__main__":
    main()
