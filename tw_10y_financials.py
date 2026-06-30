# -*- coding: utf-8 -*-
"""
台股 N 檔 10 年財務數據 tw_10y_financials.py
=======================================================================
從 FinMind 抓 watchlist 每檔近 10 年:
  - 營收 (年加總月營收, 較準)
  - 淨利 (季加總 IncomeAfterTaxes)
  - 自由現金流 (年現金流量表 OCF + CapEx)
  - 研發 (季加總 ResearchAndDevelopmentExpenses)
  - 庫存 (年底 Inventories)
  - 負債比 (年底 TotalLiabilities/TotalAssets)

Watchlist 來源(優先順序):
  1. 環境變數 TICKERS (空白/逗號/換行分隔, 適用 workflow_dispatch)
  2. data/watchlist_tw.txt
  3. 內建 fallback

輸出 data/台股_10年財務.xlsx, 8 個分頁:
  - 概覽 (每檔最新年 + 10y CAGR + merge 名稱/評等)
  - 6 個指標各一個橫向 10 年表
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股_10年財務.xlsx"
WATCHLIST_FILE = "data/watchlist_tw.txt"
WORKERS = int(os.environ.get("WORKERS", "4"))

START_YEAR = int(os.environ.get("START_YEAR", "2016"))
END_YEAR = int(os.environ.get("END_YEAR", str(datetime.now().year - 1)))
START_DATE = f"{START_YEAR}-01-01"
END_DATE = f"{END_YEAR}-12-31"


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
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
                toks.extend(t.strip() for t in line.split() if t.strip())
        if toks:
            print(f"  watchlist 來源: {WATCHLIST_FILE} ({len(toks)} 檔)")
            return list(dict.fromkeys(toks))
    fb = "2330 2454 2317 2308 0050".split()
    print(f"  watchlist 來源: 內建 fallback ({len(fb)} 檔)")
    return fb


def fm_get(dataset, sid):
    """FinMind v4 同步抓資料,回傳 DataFrame"""
    params = {"dataset": dataset, "data_id": sid,
              "start_date": START_DATE, "end_date": END_DATE}
    if TOKEN: params["token"] = TOKEN
    for attempt in range(3):
        try:
            r = requests.get(BASE, params=params, timeout=30)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429:
            time.sleep(3 * (attempt + 1)); continue
        if r.status_code != 200:
            return pd.DataFrame()
        try:
            j = r.json()
            data = j.get("data", [])
            return pd.DataFrame(data)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def fetch_one(sid):
    """抓單一公司 10 年三大報表 + 月營收"""
    try:
        mrev = fm_get("TaiwanStockMonthRevenue", sid)
        inc  = fm_get("TaiwanStockFinancialStatements", sid)
        bal  = fm_get("TaiwanStockBalanceSheet", sid)
        cf   = fm_get("TaiwanStockCashFlowsStatement", sid)
        if mrev.empty and inc.empty:
            return sid, None

        out = {}  # {year: {metric: value}}

        # === 營收: 年加總月營收(精準, 沒被四季加總誤差影響)===
        # 一律要求 12 個月才採用, 避免不完整年混進來
        if not mrev.empty and "revenue" in mrev.columns:
            mrev["year"] = mrev["date"].astype(str).str[:4]
            for y, sub in mrev.groupby("year"):
                if len(sub) < 12: continue
                out.setdefault(y, {})["營收"] = sub["revenue"].sum()

        # === 淨利 / 研發: 季加總 ===
        if not inc.empty and "type" in inc.columns:
            inc["year"] = inc["date"].astype(str).str[:4]
            for metric_key, col_aliases in [
                ("淨利", ["IncomeAfterTaxes", "ProfitAfterTax", "NetIncome", "IncomeAttributableToOwnersOfParent"]),
                ("研發", ["ResearchAndDevelopmentExpenses", "RAndDExpenses"]),
            ]:
                for alias in col_aliases:
                    sub = inc[inc["type"] == alias]
                    if sub.empty: continue
                    for y, g in sub.groupby("year"):
                        v = g["value"].sum()
                        out.setdefault(y, {}).setdefault(metric_key, v)
                    break  # 用第一個成功的 alias 就好

        # === 庫存 / 負債比: 年底快照 ===
        if not bal.empty and "type" in bal.columns:
            bal["year"] = bal["date"].astype(str).str[:4]
            # 庫存: 取最大 date(年底)
            for alias in ["Inventories", "Inventory"]:
                sub = bal[bal["type"] == alias]
                if sub.empty: continue
                for y, g in sub.groupby("year"):
                    out.setdefault(y, {})["庫存"] = g.sort_values("date").iloc[-1]["value"]
                break
            # 總負債 / 總資產
            ta_sub = bal[bal["type"] == "TotalAssets"]
            tl_sub = bal[bal["type"] == "Liabilities"]  # FinMind 是 Liabilities
            if tl_sub.empty:
                tl_sub = bal[bal["type"] == "TotalLiabilities"]
            for y in set(ta_sub["year"]) & set(tl_sub["year"]):
                ta_v = ta_sub[ta_sub["year"]==y].sort_values("date").iloc[-1]["value"]
                tl_v = tl_sub[tl_sub["year"]==y].sort_values("date").iloc[-1]["value"]
                if ta_v and ta_v > 0:
                    out.setdefault(y, {})["負債比%"] = round(tl_v / ta_v * 100, 1)

        # === 自由現金流: 年底現金流量表(YTD 累計, Q4 = 全年)===
        if not cf.empty and "type" in cf.columns:
            cf["year"] = cf["date"].astype(str).str[:4]
            # OCF 別名很多, 試多個
            ocf_aliases = ["CashFlowsFromOperatingActivities",
                           "CashFlowFromOperatingActivities",
                           "NetCashGeneratedFromOperatingActivities",
                           "CashFlowsFromOperating"]
            capex_aliases = ["PropertyPlantAndEquipment",
                             "AcquisitionOfPropertyPlantAndEquipment",
                             "CapEx",
                             "PaymentsForPropertyPlantAndEquipment"]
            ocf_dict = {}
            for a in ocf_aliases:
                sub = cf[cf["type"] == a]
                if sub.empty: continue
                # 取每年最大 date(Q4 = 全年累計)
                for y, g in sub.groupby("year"):
                    ocf_dict[y] = g.sort_values("date").iloc[-1]["value"]
                break
            capex_dict = {}
            for a in capex_aliases:
                sub = cf[cf["type"] == a]
                if sub.empty: continue
                for y, g in sub.groupby("year"):
                    capex_dict[y] = g.sort_values("date").iloc[-1]["value"]
                break
            for y, ocf in ocf_dict.items():
                cap = capex_dict.get(y, 0) or 0
                # CapEx 在 FinMind 多半為負(投入), 直接相加即 FCF
                out.setdefault(y, {})["自由現金流"] = ocf + cap

        return sid, out
    except Exception as e:
        return sid, None


def to_billions(v):
    """轉億元(台股財報是元為單位, /1e8 = 億)"""
    if v is None or pd.isna(v): return None
    try:
        return round(float(v) / 1e8, 1)
    except:
        return None


def cagr(start, end, n):
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
    if not TOKEN:
        print("⚠️ 未設 FINMIND_TOKEN(免費版速率慢, 建議付費版)")
    codes = load_watchlist()
    print(f"年份範圍: {START_YEAR} ~ {END_YEAR}")
    print(f"抓 {len(codes)} 檔財務(平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sid, data = fut.result()
            if data: results[sid] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # 從體檢總表撈名稱 / 評等 / 品質
    base = pd.DataFrame()
    for src in ["data/台股體檢總表.xlsx", "data/台股_體檢總表.xlsx"]:
        if not os.path.exists(src): continue
        try:
            h = pd.read_excel(src, sheet_name="體檢總表")
            h["代號"] = h["代號"].astype(str)
            keep = [c for c in ["代號","名稱","產業","評等","品質總分"] if c in h.columns]
            base = h[keep]
            print(f"  載入體檢總表: {src} ({len(base)} 筆)")
            break
        except Exception as e:
            print(f"  ⚠️ 讀 {src} 失敗 {e}")

    years_asc = [str(y) for y in range(START_YEAR, END_YEAR + 1)]

    metrics = ["營收","淨利","自由現金流","研發","庫存","負債比%"]
    sheets = {m: [] for m in metrics}
    overview = []

    for sid in codes:
        if sid not in results: continue
        data = results[sid]
        for m in metrics:
            row = {"代號": sid}
            for y in years_asc:
                v = data.get(y, {}).get(m)
                if m == "負債比%":
                    row[y] = v
                else:
                    row[y] = to_billions(v)
            sheets[m].append(row)

        def get_series(metric):
            return {y: data.get(y, {}).get(metric) for y in years_asc}
        rev = get_series("營收"); ni = get_series("淨利"); fcf = get_series("自由現金流")

        Y_end = str(END_YEAR)
        def gc(d, n): return cagr(d.get(str(END_YEAR - n)), d.get(Y_end), n)
        def y1(d):    return yoy(d.get(str(END_YEAR - 1)), d.get(Y_end))

        rev_e = rev.get(Y_end); ni_e = ni.get(Y_end); fcf_e = fcf.get(Y_end)
        rev_s = rev.get(str(START_YEAR)); ni_s = ni.get(str(START_YEAR))
        nm_end = round(ni_e/rev_e*100, 1) if rev_e and ni_e and rev_e > 0 else None
        nm_start = round(ni_s/rev_s*100, 1) if rev_s and ni_s and rev_s > 0 else None
        nm_delta = round(nm_end - nm_start, 1) if nm_end is not None and nm_start is not None else None
        fc_ratio = round(fcf_e/ni_e*100, 0) if fcf_e and ni_e and ni_e > 0 else None

        overview.append({
            "代號": sid,
            "起年": START_YEAR, "迄年": END_YEAR,
            f"{START_YEAR}營收(億)": to_billions(rev_s),
            f"{END_YEAR}營收(億)": to_billions(rev_e),
            f"{END_YEAR}淨利(億)": to_billions(ni_e),
            f"{END_YEAR}FCF(億)": to_billions(fcf_e),
            f"{END_YEAR}研發(億)": to_billions(data.get(Y_end, {}).get("研發")),
            f"{END_YEAR}庫存(億)": to_billions(data.get(Y_end, {}).get("庫存")),
            f"{END_YEAR}負債比%": data.get(Y_end, {}).get("負債比%"),
            "營收10y%": gc(rev, 10), "營收5y%": gc(rev, 5), "營收3y%": gc(rev, 3), "營收1y%": y1(rev),
            "淨利10y%": gc(ni, 10),  "淨利5y%": gc(ni, 5),  "淨利3y%": gc(ni, 3),  "淨利1y%": y1(ni),
            "FCF10y%":  gc(fcf, 10), "FCF5y%":  gc(fcf, 5), "FCF3y%":  gc(fcf, 3), "FCF1y%":  y1(fcf),
            "淨利率%": nm_end, "淨利率Δpp": nm_delta, "FCF/NI%": fc_ratio,
        })

    ov = pd.DataFrame(overview)
    if not base.empty:
        # 統一型別:base 代號可能是 str, results sid 是 str → ov 代號也要 str
        ov["代號"] = ov["代號"].astype(str)
        ov = ov.merge(base, on="代號", how="left")
        front = [c for c in ["代號","名稱","產業","評等","品質總分"] if c in ov.columns]
        rest = [c for c in ov.columns if c not in front]
        ov = ov[front + rest]
    ov = ov.sort_values("營收3y%", ascending=False, na_position="last")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        ov.to_excel(xw, sheet_name="概覽", index=False)
        for m in metrics:
            df = pd.DataFrame(sheets[m])
            df["代號"] = df["代號"].astype(str)
            if not base.empty and "名稱" in base.columns:
                df = df.merge(base[["代號","名稱"]], on="代號", how="left")
            cols = ["代號","名稱"] + [y for y in years_asc if y in df.columns]
            df = df[[c for c in cols if c in df.columns]]
            df.to_excel(xw, sheet_name=m, index=False)

    print(f"\n→ 已輸出 {DST}")
    print(f"分頁: 概覽 + {' / '.join(metrics)}")
    print(f"\n=== 營收 3y CAGR TOP 15 ({START_YEAR}~{END_YEAR}) ===")
    show_cols = [c for c in ["代號","名稱","評等","營收5y%","營收3y%","營收1y%","淨利3y%","FCF3y%","淨利率Δpp","FCF/NI%"] if c in ov.columns]
    print(ov[show_cols].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
