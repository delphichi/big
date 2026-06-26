# -*- coding: utf-8 -*-
"""
美股負債比補丁 patch_us_debt.py
=======================================================================
原 fetch_fundamentals_us 用的 FMP debtRatio 鍵已被 FMP stable 移除(剩 netDebtToEBITDA)
→ 體檢表「負債比%」全 NaN → ⑪短期償債警報失靈。
本腳本不重跑整個 5-endpoint 體檢(太貴),只用 ratios 端點(1 call/檔)補負債比欄。

讀:data/美股體檢總表.xlsx 體檢總表分頁
抓:每檔 1 次 ratios annual(取 debtRatio / debtToEquityRatio)
寫:更新「負債比%」欄;其餘欄位不動。
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
SRC = "data/美股體檢總表.xlsx"
WORKERS = 8
RATE_SLEEP = 0.1


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


def fetch_debt(sym):
    try:
        ratios = get("ratios", symbol=sym, period="annual", limit=3)
        if not ratios:
            return sym, None
        r0 = ratios[-1] if isinstance(ratios, list) else ratios
        debt = pick(r0, "debtRatio", "debtToAssetsRatio",
                        "totalDebtToAssets", "totalDebtToTotalAssets")
        if debt is None:
            de = pick(r0, "debtEquityRatio", "debtToEquityRatio")
            if de is not None and de > 0:
                debt = de / (1 + de)
        return sym, round(debt * 100, 1) if debt is not None else None
    except Exception as e:
        return sym, None


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY"); return
    # 讀全部 sheet,只改體檢總表的 負債比% 欄
    xls = pd.ExcelFile(SRC)
    sheets = {sh: pd.read_excel(SRC, sheet_name=sh) for sh in xls.sheet_names}
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    symbols = [s for s in h["代號"].tolist() if str(s).isalpha() and len(str(s)) <= 5]
    print(f"待補負債比: {len(symbols)} 檔(平行 {WORKERS})")
    t0 = time.time()
    results = {}
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_debt, s): s for s in symbols}
        for fut in as_completed(futs):
            try:
                sym, debt = fut.result()
            except Exception:
                continue
            results[sym] = debt
            done += 1
            if done % 200 == 0:
                ok = sum(1 for v in results.values() if v is not None)
                el = (time.time() - t0) / 60
                print(f"  [{done}/{len(symbols)}] 抓到 {ok}, 耗時 {el:.1f} 分")

    # 寫回
    h["負債比%"] = h["代號"].map(results)
    ok_cnt = h["負債比%"].notna().sum()
    print(f"\n完成:{ok_cnt}/{len(h)} 檔抓到負債比")
    # 也對其他分頁(A級好公司/A級+好價格/循環股)同步
    for sh_name, df in sheets.items():
        if sh_name == "體檢總表":
            sheets[sh_name] = h
        elif "代號" in df.columns and "負債比%" in df.columns:
            df["代號"] = df["代號"].astype(str)
            df["負債比%"] = df["代號"].map(results)
    tmp = SRC + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for sh, df in sheets.items():
            df.to_excel(xw, sheet_name=sh, index=False)
    os.replace(tmp, SRC)
    print(f"→ 已更新 {SRC}")


if __name__ == "__main__":
    main()
