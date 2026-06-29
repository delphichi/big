# -*- coding: utf-8 -*-
"""
美股 64 檔 10 年財務數據 us_10y_financials.py
=======================================================================
從 FMP 抓 watchlist 每檔近 10 年的:
  - 營收 revenue
  - 淨利 netIncome
  - 自由現金流 freeCashFlow
  - 研發支出 R&D
  - 庫存 inventory
  - 負債比(總負債/總資產)

輸出 data/美股64檔_10年財務.xlsx,8 個分頁:
  - 概覽(每檔最新年 + 10y CAGR)
  - 6 個指標各一個橫向 10 年表
  - 各檔評等/品質(merge 自體檢總表)
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
DST = "data/美股64檔_10年財務.xlsx"
WORKERS = int(os.environ.get("WORKERS", "6"))

# 同 us_pe_monitor.py 的 watchlist
WATCH = """
NVDA NVMI ANET KLAC AVGO FTNT ASML TSM LLY BRK-B MSFT META NFLX
WPM AXP AMZN CDNS CAT HWM AAPL APH GLW GOOG EXEL
HG CCEP CF AER LIN AMG COST ECL WMT FSLR IDCC NBIX
CLS LRCX AMD MU MRVL LITE AMAT CIEN COHR WDC VRT PLTR
CRM ADBE INTU NOW
GE MCO SPGI CP FER CNI
MA ORCL CHKP IDXX PAC
""".split()


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429: time.sleep(2 * (attempt+1)); continue
        if r.status_code != 200: return None
        try: return r.json()
        except: return None
    return None


def fetch_one(sym):
    """抓單一公司 10 年三大報表"""
    try:
        inc = get("income-statement", symbol=sym, period="annual", limit=10) or []
        cf = get("cash-flow-statement", symbol=sym, period="annual", limit=10) or []
        bs = get("balance-sheet-statement", symbol=sym, period="annual", limit=10) or []
        if not inc: return sym, None
        # 用 calendarYear 為 key 合併
        out = {}
        for r in inc:
            y = r.get("calendarYear") or r.get("date","")[:4]
            if not y: continue
            out.setdefault(y, {})
            out[y]["營收"] = r.get("revenue")
            out[y]["淨利"] = r.get("netIncome")
            out[y]["研發"] = r.get("researchAndDevelopmentExpenses")
        for r in cf:
            y = r.get("calendarYear") or r.get("date","")[:4]
            if not y: continue
            out.setdefault(y, {})
            out[y]["自由現金流"] = r.get("freeCashFlow")
        for r in bs:
            y = r.get("calendarYear") or r.get("date","")[:4]
            if not y: continue
            out.setdefault(y, {})
            out[y]["庫存"] = r.get("inventory")
            ta = r.get("totalAssets")
            tl = r.get("totalLiabilities") or r.get("totalDebt")
            if ta and tl:
                out[y]["負債比%"] = round(tl/ta * 100, 1)
        return sym, out
    except Exception as e:
        return sym, None


def to_billions(v):
    if v is None or pd.isna(v): return None
    try:
        return round(float(v) / 1e8, 1)  # 億美元
    except:
        return None


def main():
    if not KEY: print("⚠️ 未設 FMP_API_KEY"); return
    codes = list(dict.fromkeys(WATCH))
    print(f"抓 {len(codes)} 檔 10 年財務(平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sym, data = fut.result()
            if data: results[sym] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # 從體檢總表撈 名稱 / 評等 / 品質
    base = pd.DataFrame()
    try:
        h = pd.read_excel("data/美股體檢總表.xlsx", sheet_name="體檢總表")
        h["代號"] = h["代號"].astype(str)
        base = h[["代號","名稱","產業","評等","品質總分"]]
    except Exception as e:
        print(f"⚠️ 讀體檢總表失敗 {e}")

    # 取最近 10 年(假設 2017-2026)
    all_years = sorted({y for v in results.values() for y in v.keys()}, reverse=True)[:10]
    all_years_asc = sorted(all_years)

    # 各指標一個分頁(橫向 10 年)
    metrics = ["營收","淨利","自由現金流","研發","庫存","負債比%"]
    sheets = {m: [] for m in metrics}

    overview = []  # 概覽:每檔最新年 + CAGR

    for sym in codes:
        if sym not in results: continue
        data = results[sym]
        # 為每個指標建一行
        for m in metrics:
            row = {"代號": sym}
            for y in all_years_asc:
                v = data.get(y, {}).get(m)
                if m == "負債比%":
                    row[y] = v  # 已是百分比
                else:
                    row[y] = to_billions(v)  # 億美元
            sheets[m].append(row)

        # 概覽:最新一年 + 10y CAGR
        latest_y = max((y for y in all_years_asc if y in data), default=None)
        oldest_y = min((y for y in all_years_asc if y in data and data[y].get("營收")), default=None)
        rev_latest = data.get(latest_y, {}).get("營收") if latest_y else None
        rev_oldest = data.get(oldest_y, {}).get("營收") if oldest_y else None
        rev_cagr = None
        if rev_latest and rev_oldest and rev_oldest > 0:
            n = int(latest_y) - int(oldest_y)
            if n > 0:
                rev_cagr = round(((rev_latest/rev_oldest)**(1/n) - 1) * 100, 1)
        overview.append({
            "代號": sym,
            "最新年": latest_y,
            "營收(億)": to_billions(rev_latest),
            "淨利(億)": to_billions(data.get(latest_y, {}).get("淨利")),
            "FCF(億)": to_billions(data.get(latest_y, {}).get("自由現金流")),
            "研發(億)": to_billions(data.get(latest_y, {}).get("研發")),
            "庫存(億)": to_billions(data.get(latest_y, {}).get("庫存")),
            "負債比%": data.get(latest_y, {}).get("負債比%"),
            "10y營收CAGR%": rev_cagr,
        })

    ov = pd.DataFrame(overview)
    if not base.empty:
        ov = ov.merge(base, on="代號", how="left")
        # 重排欄順
        front = ["代號","名稱","產業","評等","品質總分"]
        rest = [c for c in ov.columns if c not in front]
        ov = ov[front + rest]
    ov = ov.sort_values("10y營收CAGR%", ascending=False, na_position="last")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        ov.to_excel(xw, sheet_name="概覽", index=False)
        for m in metrics:
            df = pd.DataFrame(sheets[m])
            if not base.empty: df = df.merge(base[["代號","名稱"]], on="代號", how="left")
            cols = ["代號","名稱"] + sorted([c for c in df.columns if c not in ("代號","名稱")])
            df = df[[c for c in cols if c in df.columns]]
            df.to_excel(xw, sheet_name=m, index=False)

    print(f"\n→ 已輸出 {DST}")
    print(f"分頁:概覽 + {' / '.join(metrics)}")
    print(f"\n=== 10y 營收 CAGR TOP 15 ===")
    print(ov[["代號","名稱","評等","營收(億)","FCF(億)","10y營收CAGR%"]].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
