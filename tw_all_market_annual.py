# -*- coding: utf-8 -*-
"""
台股全市場年度數據 tw_all_market_annual.py
=======================================================================
用 FinMind bulk 模式(不帶 data_id) 一次抓全市場 10 年:
  - 營收(月營收 12 個月加總)
  - 淨利(季損益 IncomeAfterTaxes 加總)
  - 自由現金流(季 OCF + CapEx 年底 YTD)
  - 負債比(季 BS 年底快照)
  - 每年第一交易日收盤價(TaiwanStockPriceAdj bulk)
  - 每年最後交易日收盤價
  - 年報酬率(1y/3y/5y/10y CAGR)

Bulk API calls: ~180 calls
  - 40 季 × 3 表 (INC/BS/CF) = 120 calls
  - 120 個月 monthly revenue (chunked by 季) = 40 calls
  - 20 個交易日 price snapshot = 20 calls
  - 1 call 抓股票 info

估計 30-60 分鐘 (FinMind 600/hr)

輸出 data/台股全市場_年度數據.xlsx:
  - 總覽 (每檔 × 每年一欄, 主要指標)
  - 營收 / 淨利 / FCF / 負債比 / 年頭價 / 年尾價 / 年報酬% (各分頁)
  - 成長率總表 (YoY 各年 + 10y/5y/3y CAGR)
"""
import os, time, requests, sys
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股全市場_年度數據.xlsx"
WORKERS = int(os.environ.get("WORKERS", "4"))

START_YEAR = int(os.environ.get("START_YEAR", "2016"))
END_YEAR = int(os.environ.get("END_YEAR", "2025"))


def fm(dataset, data_id=None, start=None, end=None, retry=5):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for i in range(retry):
        try:
            r = requests.get(BASE, params=p, timeout=90)
            if r.status_code == 402:
                print(f"    ⚠️ 402 付費限制: {dataset}"); return pd.DataFrame()
            if r.status_code == 429:
                wait = 5 * (i+1); print(f"    429, wait {wait}s..."); time.sleep(wait); continue
            if r.status_code != 200:
                print(f"    HTTP {r.status_code}: {dataset}"); return pd.DataFrame()
            j = r.json()
            return pd.DataFrame(j.get("data", []))
        except Exception as e:
            print(f"    {dataset} err: {e}"); time.sleep(2)
    return pd.DataFrame()


def quarters_between(sy, ey):
    """回傳 [(year, quarter_num, start_date, end_date), ...] for 每季"""
    out = []
    for y in range(sy, ey + 1):
        for q, (m1, m2) in enumerate([(1,3),(4,6),(7,9),(10,12)], 1):
            end_day = 31 if m2 in (1,3,5,7,8,10,12) else (30 if m2 in (4,6,9,11) else 28)
            out.append((y, q, f"{y}-{m1:02d}-01", f"{y}-{m2:02d}-{end_day}"))
    return out


def fetch_quarter(dataset, year, quarter, start, end):
    """抓某季全市場 dataset"""
    df = fm(dataset, start=start, end=end)
    return dataset, year, quarter, df


def get_trading_days_bounds(year):
    """回傳該年第一/最後交易日 (簡單用 1/1-1/15, 12/15-12/31 抓區間)"""
    return (f"{year}-01-01", f"{year}-01-15"), (f"{year}-12-15", f"{year}-12-31")


def main():
    if not TOKEN:
        print("⚠️ 需 FinMind Backer/Sponsor token"); sys.exit(1)

    print(f"=== 台股全市場年度數據 {START_YEAR}~{END_YEAR} ===\n")

    # ─── 1. 抓所有股票代號 ───
    print("1. 抓 TaiwanStockInfo...")
    info = fm("TaiwanStockInfo")
    if info.empty:
        print("⚠️ 無法抓股票 info"); sys.exit(1)
    # 只留普通股(排除 ETF/權證/受益證券, 常見規則:5 碼以下 + 非數字結尾)
    info["stock_id"] = info["stock_id"].astype(str)
    common = info[info["stock_id"].str.match(r"^\d{4,5}$") & ~info["stock_id"].str.startswith("00")]
    common = common.drop_duplicates(subset="stock_id")
    print(f"   → {len(common)} 檔普通股")
    stock_map = dict(zip(common["stock_id"], common["stock_name"]))

    # ─── 2. Bulk 抓 季損益 / 季 BS / 季 CF ───
    quarters = quarters_between(START_YEAR, END_YEAR)
    print(f"\n2. 抓 40 季 × 3 表 = 120 calls (平行 {WORKERS})...")
    all_inc, all_bs, all_cf = [], [], []
    jobs = []
    for y, q, s, e in quarters:
        for dataset, bucket in [
            ("TaiwanStockFinancialStatements", all_inc),
            ("TaiwanStockBalanceSheet", all_bs),
            ("TaiwanStockCashFlowsStatement", all_cf),
        ]:
            jobs.append((dataset, y, q, s, e, bucket))
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_quarter, ds, y, q, s, e): (ds, y, q, bucket) for ds, y, q, s, e, bucket in jobs}
        done = 0
        for fut in as_completed(futs):
            ds, y, q, df = fut.result()
            done += 1
            if not df.empty:
                df = df.assign(__year=y, __quarter=q)
                for a, b, target in [
                    ("TaiwanStockFinancialStatements", ds, all_inc),
                    ("TaiwanStockBalanceSheet", ds, all_bs),
                    ("TaiwanStockCashFlowsStatement", ds, all_cf),
                ]:
                    if ds == a: target.append(df); break
            if done % 20 == 0: print(f"   [{done}/{len(jobs)}]")

    inc = pd.concat(all_inc, ignore_index=True) if all_inc else pd.DataFrame()
    bs  = pd.concat(all_bs, ignore_index=True) if all_bs else pd.DataFrame()
    cf  = pd.concat(all_cf, ignore_index=True) if all_cf else pd.DataFrame()
    print(f"   INC: {len(inc)} 筆 / BS: {len(bs)} 筆 / CF: {len(cf)} 筆")

    # ─── 3. 抓 月營收 (逐季 bulk) ───
    print(f"\n3. 抓 monthly revenue ({len(quarters)} 季 bulk)...")
    all_rev = []
    def fetch_mrev(y, q, s, e):
        df = fm("TaiwanStockMonthRevenue", start=s, end=e)
        return y, q, df
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_mrev, y, q, s, e): (y, q) for y, q, s, e in quarters}
        done = 0
        for fut in as_completed(futs):
            y, q, df = fut.result()
            done += 1
            if not df.empty:
                df["__year"] = y
                all_rev.append(df)
            if done % 10 == 0: print(f"   [{done}/{len(quarters)}]")
    rev = pd.concat(all_rev, ignore_index=True) if all_rev else pd.DataFrame()
    print(f"   月營收: {len(rev)} 筆")

    # ─── 4. 抓 每年頭尾價 (bulk PriceAdj) ───
    print(f"\n4. 抓每年頭尾交易日還原股價 ({(END_YEAR-START_YEAR+1)*2} 區間)...")
    all_price = []
    for y in range(START_YEAR, END_YEAR + 1):
        for label, (s, e) in [("head", (f"{y}-01-01", f"{y}-01-15")),
                              ("tail", (f"{y}-12-15", f"{y}-12-31"))]:
            df = fm("TaiwanStockPriceAdj", start=s, end=e)
            if not df.empty:
                df["__year"] = y; df["__which"] = label
                all_price.append(df)
    price = pd.concat(all_price, ignore_index=True) if all_price else pd.DataFrame()
    print(f"   價格: {len(price)} 筆")

    # ─── 5. 聚合成 (stock_id, year) 表 ───
    print(f"\n5. 聚合處理...")

    def pick_bs(df, want):
        """從 BS/INC/CF wide-ish long 表撈某 type"""
        if df.empty or "type" not in df.columns: return {}
        sub = df[df["type"].isin(want if isinstance(want, list) else [want])]
        return sub

    years = list(range(START_YEAR, END_YEAR + 1))

    # 損益: 淨利(4 季 sum)
    ni_agg = {}
    if not inc.empty:
        ni_types = ["IncomeAfterTaxes","ProfitAfterTax","NetIncome","IncomeAttributableToOwnersOfParent"]
        for t in ni_types:
            sub = inc[inc.get("type","") == t]
            if sub.empty: continue
            for (sid, y), g in sub.groupby(["stock_id","__year"]):
                v = g["value"].sum()
                ni_agg.setdefault((str(sid), y), v)
            if ni_agg: break

    # 現金流: OCF - CapEx (取每年 Q4 = 全年 YTD 累計)
    fcf_agg = {}
    if not cf.empty:
        ocf_alias = ["CashFlowsFromOperatingActivities","CashFlowFromOperatingActivities","NetCashGeneratedFromOperatingActivities"]
        capex_alias = ["PropertyPlantAndEquipment","AcquisitionOfPropertyPlantAndEquipment","CapEx"]
        ocf_d, capex_d = {}, {}
        for alias in ocf_alias:
            sub = cf[cf.get("type","") == alias]
            if sub.empty: continue
            for (sid, y), g in sub.groupby(["stock_id","__year"]):
                # 取 Q4 那筆 (最大 date)
                latest = g.sort_values("date").iloc[-1]
                ocf_d[(str(sid), y)] = latest["value"]
            break
        for alias in capex_alias:
            sub = cf[cf.get("type","") == alias]
            if sub.empty: continue
            for (sid, y), g in sub.groupby(["stock_id","__year"]):
                latest = g.sort_values("date").iloc[-1]
                capex_d[(str(sid), y)] = latest["value"]
            break
        for k, ocf_v in ocf_d.items():
            cap = capex_d.get(k, 0) or 0
            fcf_agg[k] = (ocf_v or 0) + cap  # CapEx 多半為負

    # 負債比: 年底 BS TotalLiabilities / TotalAssets
    dr_agg = {}
    if not bs.empty:
        ta = bs[bs.get("type","") == "TotalAssets"]
        tl = bs[bs.get("type","").isin(["Liabilities","TotalLiabilities"])]
        if not ta.empty and not tl.empty:
            ta_grp = ta.sort_values("date").groupby(["stock_id","__year"]).last()
            tl_grp = tl.sort_values("date").groupby(["stock_id","__year"]).last()
            for idx in ta_grp.index.intersection(tl_grp.index):
                tav = ta_grp.loc[idx, "value"]; tlv = tl_grp.loc[idx, "value"]
                if tav and tav > 0:
                    dr_agg[(str(idx[0]), idx[1])] = round(tlv / tav * 100, 1)

    # 年營收: 月營收 12 個月加總
    rev_agg = {}
    if not rev.empty and "revenue" in rev.columns:
        for (sid, y), g in rev.groupby(["stock_id","__year"]):
            if len(g) < 12 and y < datetime.now().year:  # 不滿 12 個月跳過
                continue
            rev_agg[(str(sid), y)] = g["revenue"].sum()

    # 年頭 / 年尾價
    head_price = {}; tail_price = {}
    if not price.empty:
        head = price[price["__which"] == "head"].sort_values("date").groupby(["stock_id","__year"]).first()
        tail = price[price["__which"] == "tail"].sort_values("date").groupby(["stock_id","__year"]).last()
        for idx, row in head.iterrows():
            head_price[(str(idx[0]), idx[1])] = row.get("close")
        for idx, row in tail.iterrows():
            tail_price[(str(idx[0]), idx[1])] = row.get("close")

    # ─── 6. 組長表 → pivot 成寬表 ───
    all_ids = set()
    for d in [ni_agg, fcf_agg, dr_agg, rev_agg, head_price, tail_price]:
        for sid, y in d.keys():
            all_ids.add(sid)
    all_ids = sorted(all_ids)
    print(f"   共 {len(all_ids)} 檔有資料")

    def build_sheet(agg, div=1e8, is_pct=False):
        rows = []
        for sid in all_ids:
            row = {"代號": sid, "名稱": stock_map.get(sid, "")}
            for y in years:
                v = agg.get((sid, y))
                if v is None or pd.isna(v):
                    row[str(y)] = None
                else:
                    row[str(y)] = round(float(v) / div, 2) if not is_pct else round(float(v), 1)
            rows.append(row)
        return pd.DataFrame(rows)

    sh_rev = build_sheet(rev_agg, div=1e8)      # 億
    sh_ni  = build_sheet(ni_agg, div=1e8)
    sh_fcf = build_sheet(fcf_agg, div=1e8)
    sh_dr  = build_sheet(dr_agg, div=1, is_pct=True)  # %
    sh_hp  = build_sheet(head_price, div=1, is_pct=True)   # 元
    sh_tp  = build_sheet(tail_price, div=1, is_pct=True)

    # 年報酬 = 年尾價 / 年頭價 - 1
    def yearly_return_sheet():
        rows = []
        for sid in all_ids:
            row = {"代號": sid, "名稱": stock_map.get(sid, "")}
            for y in years:
                h = head_price.get((sid, y)); t = tail_price.get((sid, y))
                if h and t and h > 0:
                    row[str(y)] = round((t/h - 1) * 100, 1)
                else:
                    row[str(y)] = None
            rows.append(row)
        return pd.DataFrame(rows)
    sh_ret = yearly_return_sheet()

    # 成長率總表: 各檔 10y/5y/3y 營收/淨利 CAGR
    def cagr(s, e, n):
        if s is None or e is None or pd.isna(s) or pd.isna(e) or s <= 0 or e <= 0 or n <= 0: return None
        try: return round(((e/s)**(1/n) - 1) * 100, 1)
        except: return None

    growth_rows = []
    for sid in all_ids:
        name = stock_map.get(sid, "")
        row = {"代號": sid, "名稱": name}
        for metric, agg in [("營收", rev_agg), ("淨利", ni_agg), ("FCF", fcf_agg)]:
            e = agg.get((sid, END_YEAR))
            for n in [10, 5, 3]:
                s = agg.get((sid, END_YEAR - n))
                row[f"{metric}{n}y%"] = cagr(s, e, n)
        # 10y 股價報酬 (年頭 sy → 年尾 ey)
        s_price = head_price.get((sid, START_YEAR))
        e_price = tail_price.get((sid, END_YEAR))
        if s_price and e_price and s_price > 0:
            total_ret = round((e_price/s_price - 1) * 100, 1)
            row["10y股價%"] = total_ret
            row["10y年化%"] = cagr(s_price, e_price, END_YEAR - START_YEAR)
        growth_rows.append(row)
    sh_growth = pd.DataFrame(growth_rows)

    # ─── 7. 輸出 ───
    os.makedirs("data", exist_ok=True)
    print(f"\n7. 寫入 {DST}...")
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        # 總覽 = 成長率總表 (最實用)
        sh_growth.sort_values("營收3y%", ascending=False, na_position="last").to_excel(
            xw, sheet_name="成長率總覽", index=False)
        sh_rev.to_excel(xw, sheet_name="營收(億)", index=False)
        sh_ni.to_excel(xw, sheet_name="淨利(億)", index=False)
        sh_fcf.to_excel(xw, sheet_name="FCF(億)", index=False)
        sh_dr.to_excel(xw, sheet_name="負債比%", index=False)
        sh_hp.to_excel(xw, sheet_name="年頭價", index=False)
        sh_tp.to_excel(xw, sheet_name="年尾價", index=False)
        sh_ret.to_excel(xw, sheet_name="年報酬%", index=False)

    print(f"→ {DST}")
    print(f"   {len(all_ids)} 檔 × {len(years)} 年 × 8 分頁")


if __name__ == "__main__":
    main()
