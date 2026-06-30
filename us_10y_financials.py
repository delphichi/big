# -*- coding: utf-8 -*-
"""
美股 N 檔多年財務數據 us_10y_financials.py
=======================================================================
從 FMP 抓 watchlist 每檔指定年份範圍的:
  - 營收 revenue
  - 淨利 netIncome
  - 自由現金流 freeCashFlow
  - 研發支出 R&D
  - 庫存 inventory
  - 負債比 (總負債/總資產)

Watchlist 來源(優先順序):
  1. 環境變數 TICKERS (空白/逗號/換行分隔, 適用 workflow_dispatch)
  2. data/watchlist_us.txt (一行一檔, # 開頭忽略)
  3. 內建 fallback

年份範圍:
  START_YEAR / END_YEAR (環境變數, 預設 2016 / 上一年)
  例: START_YEAR=2016 END_YEAR=2025

輸出 data/美股_10年財務.xlsx, 8 個分頁:
  - 概覽: 每檔起迄年數據 + 10y/5y/3y/1y 三大成長率 + 淨利率Δ + FCF轉換率
  - 6 個指標各一個橫向年表
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
DST = "data/美股_10年財務.xlsx"
WATCHLIST_FILE = "data/watchlist_us.txt"
WORKERS = int(os.environ.get("WORKERS", "6"))

START_YEAR = int(os.environ.get("START_YEAR", "2016"))
END_YEAR = int(os.environ.get("END_YEAR", str(datetime.now().year - 1)))
# 抓多一點 buffer, 過濾再裁切
FETCH_LIMIT = max(15, END_YEAR - START_YEAR + 5)


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip().upper() for t in env.replace(",", " ").split() if t.strip()]
        toks = [t for t in toks if t and not t.startswith("#")]
        if toks:
            print(f"  watchlist 來源: 環境變數 TICKERS ({len(toks)} 檔)")
            return list(dict.fromkeys(toks))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip().upper() for t in line.split() if t.strip())
        if toks:
            print(f"  watchlist 來源: {WATCHLIST_FILE} ({len(toks)} 檔)")
            return list(dict.fromkeys(toks))
    fb = "NVDA AVGO TSM META GOOG MSFT".split()
    print(f"  watchlist 來源: 內建 fallback ({len(fb)} 檔)")
    return fb


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


DEBUG_TICKER = os.environ.get("DEBUG_TICKER", "").upper().strip()


def fetch_one(sym):
    """抓單一公司財報, 按年份過濾"""
    try:
        inc = get("income-statement", symbol=sym, period="annual", limit=FETCH_LIMIT) or []
        cf  = get("cash-flow-statement", symbol=sym, period="annual", limit=FETCH_LIMIT) or []
        bs  = get("balance-sheet-statement", symbol=sym, period="annual", limit=FETCH_LIMIT) or []
        if not inc: return sym, None

        # === DEBUG: 印出 raw 日期欄位讓 user 對照 ===
        if sym == DEBUG_TICKER:
            print(f"\n=== DEBUG {sym} raw records ===")
            for label, recs in [("INC", inc), ("CF", cf), ("BS", bs)]:
                print(f"--- {label} ({len(recs)} 筆) ---")
                for r in recs:
                    print(f"  date={r.get('date')}  calendarYear={r.get('calendarYear')}  "
                          f"period={r.get('period')}  acceptedDate={r.get('acceptedDate')}  "
                          f"fillingDate={r.get('fillingDate')}  rev={r.get('revenue')}")
            print()
        out = {}
        for r in inc:
            # 用 date (fiscal year end) 優先, calendarYear 有些公司會錯位
            # (例: EXEL FY2024 10-K 2025/2 才送, FMP 標 calendarYear=2025 → 2024 空)
            y = r.get("date","")[:4] or r.get("calendarYear")
            if not y or not str(y).isdigit(): continue
            yi = int(y)
            if yi < START_YEAR or yi > END_YEAR: continue
            out.setdefault(y, {})
            out[y]["營收"] = r.get("revenue")
            out[y]["淨利"] = r.get("netIncome")
            out[y]["研發"] = r.get("researchAndDevelopmentExpenses")
        for r in cf:
            # 用 date (fiscal year end) 優先, calendarYear 有些公司會錯位
            # (例: EXEL FY2024 10-K 2025/2 才送, FMP 標 calendarYear=2025 → 2024 空)
            y = r.get("date","")[:4] or r.get("calendarYear")
            if not y or not str(y).isdigit(): continue
            yi = int(y)
            if yi < START_YEAR or yi > END_YEAR: continue
            out.setdefault(y, {})
            out[y]["自由現金流"] = r.get("freeCashFlow")
        for r in bs:
            # 用 date (fiscal year end) 優先, calendarYear 有些公司會錯位
            # (例: EXEL FY2024 10-K 2025/2 才送, FMP 標 calendarYear=2025 → 2024 空)
            y = r.get("date","")[:4] or r.get("calendarYear")
            if not y or not str(y).isdigit(): continue
            yi = int(y)
            if yi < START_YEAR or yi > END_YEAR: continue
            out.setdefault(y, {})
            out[y]["庫存"] = r.get("inventory")
            ta = r.get("totalAssets")
            tl = r.get("totalLiabilities") or r.get("totalDebt")
            if ta and tl:
                out[y]["負債比%"] = round(tl/ta * 100, 1)
        return sym, out
    except Exception:
        return sym, None


def to_billions(v):
    if v is None or pd.isna(v): return None
    try: return round(float(v) / 1e8, 1)
    except: return None


def cagr(start, end, n):
    """End-of-period CAGR, 起迄都要 > 0 才算"""
    if start is None or end is None: return None
    try:
        if start <= 0 or end <= 0 or n <= 0: return None
        return round(((end/start)**(1/n) - 1) * 100, 1)
    except: return None


def yoy(prev, cur):
    if prev is None or cur is None: return None
    try:
        if prev == 0: return None
        return round((cur/prev - 1) * 100, 1)
    except: return None


def main():
    if not KEY: print("⚠️ 未設 FMP_API_KEY"); return
    codes = load_watchlist()
    print(f"年份範圍: {START_YEAR} ~ {END_YEAR}")
    print(f"抓 {len(codes)} 檔財務 (平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sym, data = fut.result()
            if data: results[sym] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # 體檢總表 merge
    base = pd.DataFrame()
    try:
        h = pd.read_excel("data/美股體檢總表.xlsx", sheet_name="體檢總表")
        h["代號"] = h["代號"].astype(str)
        base = h[["代號","名稱","產業","評等","品質總分"]]
    except Exception as e:
        print(f"⚠️ 讀體檢總表失敗 {e}")

    # 年份列表(asc)
    years_asc = [str(y) for y in range(START_YEAR, END_YEAR + 1)]
    n_years = END_YEAR - START_YEAR  # 起迄之間的年數差

    metrics = ["營收","淨利","自由現金流","研發","庫存","負債比%"]
    sheets = {m: [] for m in metrics}
    overview = []

    for sym in codes:
        if sym not in results: continue
        data = results[sym]

        # 每指標一橫列
        for m in metrics:
            row = {"代號": sym}
            for y in years_asc:
                v = data.get(y, {}).get(m)
                row[y] = v if m == "負債比%" else to_billions(v)
            sheets[m].append(row)

        # 三大指標數值 (raw, 算成長率用)
        def get_series(metric):
            return {y: data.get(y, {}).get(metric) for y in years_asc}
        rev = get_series("營收")
        ni  = get_series("淨利")
        fcf = get_series("自由現金流")

        Y_end = str(END_YEAR)
        # 各期間成長率
        def gc(d, n):
            return cagr(d.get(str(END_YEAR - n)), d.get(Y_end), n)
        def y1(d):
            return yoy(d.get(str(END_YEAR - 1)), d.get(Y_end))
        # 淨利率 / FCF 轉換率
        rev_e = rev.get(Y_end); ni_e = ni.get(Y_end); fcf_e = fcf.get(Y_end)
        rev_s = rev.get(str(START_YEAR)); ni_s = ni.get(str(START_YEAR))
        nm_end = round(ni_e/rev_e*100, 1) if rev_e and ni_e and rev_e > 0 else None
        nm_start = round(ni_s/rev_s*100, 1) if rev_s and ni_s and rev_s > 0 else None
        nm_delta = round(nm_end - nm_start, 1) if nm_end is not None and nm_start is not None else None
        fc_ratio = round(fcf_e/ni_e*100, 0) if fcf_e and ni_e and ni_e > 0 else None

        overview.append({
            "代號": sym,
            "起年": START_YEAR, "迄年": END_YEAR,
            f"{START_YEAR}營收(億)": to_billions(rev_s),
            f"{END_YEAR}營收(億)": to_billions(rev_e),
            f"{END_YEAR}淨利(億)": to_billions(ni_e),
            f"{END_YEAR}FCF(億)": to_billions(fcf_e),
            f"{END_YEAR}研發(億)": to_billions(data.get(Y_end, {}).get("研發")),
            f"{END_YEAR}庫存(億)": to_billions(data.get(Y_end, {}).get("庫存")),
            f"{END_YEAR}負債比%": data.get(Y_end, {}).get("負債比%"),
            # 三大成長率 × 4 期間
            "營收10y%": gc(rev, 10), "營收5y%": gc(rev, 5), "營收3y%": gc(rev, 3), "營收1y%": y1(rev),
            "淨利10y%": gc(ni, 10),  "淨利5y%": gc(ni, 5),  "淨利3y%": gc(ni, 3),  "淨利1y%": y1(ni),
            "FCF10y%":  gc(fcf, 10), "FCF5y%":  gc(fcf, 5), "FCF3y%":  gc(fcf, 3), "FCF1y%":  y1(fcf),
            "淨利率%": nm_end, "淨利率Δpp": nm_delta, "FCF/NI%": fc_ratio,
        })

    ov = pd.DataFrame(overview)
    if not base.empty:
        ov = ov.merge(base, on="代號", how="left")
        front = ["代號","名稱","產業","評等","品質總分"]
        rest = [c for c in ov.columns if c not in front]
        ov = ov[front + rest]
    ov = ov.sort_values("營收3y%", ascending=False, na_position="last")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        ov.to_excel(xw, sheet_name="概覽", index=False)
        for m in metrics:
            df = pd.DataFrame(sheets[m])
            if not base.empty: df = df.merge(base[["代號","名稱"]], on="代號", how="left")
            cols = ["代號","名稱"] + [y for y in years_asc if y in df.columns]
            df = df[[c for c in cols if c in df.columns]]
            df.to_excel(xw, sheet_name=m, index=False)

    print(f"\n→ 已輸出 {DST}")
    print(f"分頁: 概覽 + {' / '.join(metrics)}")
    print(f"\n=== 營收 3y CAGR TOP 15 ({START_YEAR}~{END_YEAR}) ===")
    cols_show = [c for c in ["代號","名稱","評等","營收5y%","營收3y%","營收1y%","淨利3y%","FCF3y%","淨利率Δpp","FCF/NI%"] if c in ov.columns]
    print(ov[cols_show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
