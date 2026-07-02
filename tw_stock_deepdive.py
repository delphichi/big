# -*- coding: utf-8 -*-
"""
台股單檔深度研究 tw_stock_deepdive.py
=======================================================================
對單一台股跑 ~25 個 FinMind dataset, 整合完整研究報告

抓內容:
  1. 公司基本 (TaiwanStockInfo)
  2. 產業鏈 (TaiwanStockIndustryChain)
  3. 還原股價 1y/10y + PER/PBR 歷史 3y
  4. 月營收 10Y (YoY / MoM)
  5. 季損益/資產負債/現金流 10Y
  6. 市值歷史
  7. 三大法人 90d + 外資持股歷史
  8. 八大行庫 / 融資融券 / 借券 / 股權分級
  9. 股利歷史 + 除權除息
  10. 拆股 / 減資
  11. 警戒 (處置/暫停/融券暫停)
  12. 個股新聞
  13. 產業鏈同業

跑法:
  FINMIND_TOKEN=xxx python tw_stock_deepdive.py 2330
  FINMIND_TOKEN=xxx python tw_stock_deepdive.py 2330 6139 5274
  TICKER=2330 python tw_stock_deepdive.py

輸出:
  data/deepdive/台股深度_{代號}.xlsx  (~20 分頁完整資料)
  data/deepdive/台股深度_{代號}.md    (Markdown 精華)
"""
import os, sys, time, requests
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
OUT_DIR = os.environ.get("OUT_DIR", "data/deepdive")
END = datetime.now().strftime("%Y-%m-%d")
D90 = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
D1Y = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
D3Y = (datetime.now() - timedelta(days=365*3)).strftime("%Y-%m-%d")
D10Y = (datetime.now() - timedelta(days=365*10)).strftime("%Y-%m-%d")


def fm(dataset, data_id=None, start=None, end=None):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for i in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=30)
            if r.status_code == 402: return pd.DataFrame()
            if r.status_code == 429: time.sleep(3*(i+1)); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception: time.sleep(1)
    return pd.DataFrame()


def fetch_all(sid):
    tasks = [
        ("info",         ("TaiwanStockInfo", None, None, None)),
        ("industry",     ("TaiwanStockIndustryChain", None, None, None)),
        ("price_1y",     ("TaiwanStockPriceAdj", sid, D1Y, END)),
        ("price_10y",    ("TaiwanStockPriceAdj", sid, D10Y, END)),
        ("per",          ("TaiwanStockPER", sid, D3Y, END)),
        ("month_rev",    ("TaiwanStockMonthRevenue", sid, D10Y, END)),
        ("income",       ("TaiwanStockFinancialStatements", sid, D10Y, END)),
        ("bs",           ("TaiwanStockBalanceSheet", sid, D10Y, END)),
        ("cf",           ("TaiwanStockCashFlowsStatement", sid, D10Y, END)),
        ("market_val",   ("TaiwanStockMarketValue", sid, D3Y, END)),
        ("inst_wide",    ("TaiwanStockInstitutionalInvestorsBuySellWide", sid, D90, END)),
        ("shareholding", ("TaiwanStockShareholding", sid, D1Y, END)),
        ("gov_bank",     ("TaiwanstockGovernmentBankBuySell", sid, D90, END)),
        ("margin",       ("TaiwanStockMarginPurchaseShortSale", sid, D1Y, END)),
        ("lending",      ("TaiwanStockSecuritiesLending", sid, D90, END)),
        ("shares_per",   ("TaiwanStockHoldingSharesPer", sid, D1Y, END)),
        ("dividend",     ("TaiwanStockDividend", sid, D10Y, END)),
        ("div_result",   ("TaiwanStockDividendResult", sid, D3Y, END)),
        ("split",        ("TaiwanStockSplitPrice", sid, D10Y, END)),
        ("capred",       ("TaiwanStockCapitalReductionReferencePrice", sid, D10Y, END)),
        ("disp",         ("TaiwanStockDispositionSecuritiesPeriod", sid, D1Y, END)),
        ("suspend",      ("TaiwanStockSuspended", sid, D1Y, END)),
        ("ms_suspend",   ("TaiwanStockMarginShortSaleSuspension", sid, D1Y, END)),
        ("news",         ("TaiwanStockNews", sid, D90, END)),
    ]
    out = {}
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(fm, *args): key for key, args in tasks}
        for fut in as_completed(futs):
            key = futs[fut]
            try: out[key] = fut.result()
            except Exception: out[key] = pd.DataFrame()
    return out


def build_xlsx(sid, name, data, dst):
    sheets = {}

    # 產業
    ind_dataset = data.get("industry", pd.DataFrame())
    industry_str = ""
    if not ind_dataset.empty and "stock_id" in ind_dataset.columns:
        my = ind_dataset[ind_dataset["stock_id"].astype(str) == str(sid)]
        if len(my) and "industry_category" in my.columns:
            industry_str = " / ".join(my["industry_category"].astype(str).unique()[:3])

    # 股價
    price_1y = data.get("price_1y", pd.DataFrame())
    latest_price = year_high = year_low = None
    if not price_1y.empty and "close" in price_1y.columns:
        p = price_1y.sort_values("date")
        latest_price = float(p.iloc[-1]["close"])
        year_high = float(p["max"].max()) if "max" in p.columns else float(p["close"].max())
        year_low = float(p["min"].min()) if "min" in p.columns else float(p["close"].min())

    # PER/PBR
    per = data.get("per", pd.DataFrame())
    latest_per = latest_pbr = div_yield = None
    if not per.empty:
        per_s = per.sort_values("date")
        latest_per = per_s.iloc[-1].get("PER")
        latest_pbr = per_s.iloc[-1].get("PBR")
        div_yield = per_s.iloc[-1].get("dividend_yield")

    # 市值
    mv = data.get("market_val", pd.DataFrame())
    latest_mv = None
    if not mv.empty and "market_value" in mv.columns:
        latest_mv = mv.sort_values("date").iloc[-1]["market_value"]

    # 月營收 YoY (FinMind 常沒回 growth 欄, 自己算)
    mrev = data.get("month_rev", pd.DataFrame())
    rev_yoy = latest_month_rev = None
    if not mrev.empty and "revenue" in mrev.columns:
        mrev_s = mrev.sort_values("date").reset_index(drop=True)
        latest_month_rev = mrev_s.iloc[-1].get("revenue")
        if len(mrev_s) >= 13:
            cur = float(mrev_s.iloc[-1]["revenue"] or 0)
            prev = float(mrev_s.iloc[-13]["revenue"] or 0)
            if prev > 0:
                rev_yoy = round((cur / prev - 1) * 100, 1)
        if rev_yoy is None and "revenue_year_growth" in mrev.columns:
            rev_yoy = mrev_s.iloc[-1].get("revenue_year_growth")

    # 三大法人 90d
    iw = data.get("inst_wide", pd.DataFrame())
    def s(df, col): return int(df[col].sum()) if col in df.columns else 0
    f_net = t_net = d_net = 0
    if not iw.empty:
        f_net = s(iw,"Foreign_Investor_buy") - s(iw,"Foreign_Investor_sell")
        t_net = s(iw,"Investment_Trust_buy") - s(iw,"Investment_Trust_sell")
        d_net = (s(iw,"Dealer_buy")+s(iw,"Dealer_self_buy")+s(iw,"Dealer_Hedging_buy")
                 - s(iw,"Dealer_sell")-s(iw,"Dealer_self_sell")-s(iw,"Dealer_Hedging_sell"))

    # 外資持股
    sh = data.get("shareholding", pd.DataFrame())
    foreign_pct = None
    if not sh.empty and "ForeignInvestmentSharesRatio" in sh.columns:
        foreign_pct = sh.sort_values("date").iloc[-1]["ForeignInvestmentSharesRatio"]

    # 融資融券
    mg = data.get("margin", pd.DataFrame())
    margin_balance = short_balance = None
    if not mg.empty:
        mg_s = mg.sort_values("date")
        if "MarginPurchaseTodayBalance" in mg_s.columns:
            margin_balance = int(mg_s.iloc[-1]["MarginPurchaseTodayBalance"])
        if "ShortSaleTodayBalance" in mg_s.columns:
            short_balance = int(mg_s.iloc[-1]["ShortSaleTodayBalance"])

    # 警戒
    disp_count = len(data.get("disp", pd.DataFrame()))
    suspend_count = len(data.get("suspend", pd.DataFrame()))

    # 3:1 入場 + 事件雷達 (供概覽用)
    entry = _calc_entry_tw(latest_price, year_low, year_high)
    trend = _foreign_trend_tw(iw, sh)

    # PER 換算 (供概覽用)
    _inc = data.get("income", pd.DataFrame())
    per_md, per_df = _per_band_tw(latest_price, _inc)
    per_summary = None
    if per_df is not None and not per_df.empty:
        eps_4q_rows = _inc[_inc["type"] == "EPS"].sort_values("date", ascending=False).head(4)
        eps_4q = round(float(eps_4q_rows["value"].sum()), 2)
        per_summary = {"eps_4q": eps_4q,
                       "cur_pe": round(latest_price / eps_4q, 1) if eps_4q > 0 else None,
                       "fair_22x": round(eps_4q * 22, 2) if eps_4q > 0 else None,
                       "conservative_18x": round(eps_4q * 18, 2) if eps_4q > 0 else None}

    # === 概覽 ===
    ov = [
        ("代號", sid), ("名稱", name), ("產業", industry_str),
        ("─── 報價 ───", ""),
        ("當前價", latest_price),
        ("52w 高", year_high), ("52w 低", year_low),
        ("市值(億)", round((latest_mv or 0)/1e8, 1) if latest_mv else None),
        ("─── 估值 ───", ""),
        ("PER", latest_per), ("PBR", latest_pbr), ("殖利率%", div_yield),
        ("─── 動能 ───", ""),
        ("最新月營收(億)", round((latest_month_rev or 0)/1e8, 1) if latest_month_rev else None),
        ("月營收 YoY%", rev_yoy),
        ("─── 3:1 入場 ───", ""),
        ("3:1 判讀", entry["verdict"] if entry else None),
        ("SL 止損", entry["sl"] if entry else None),
        ("TP 目標", entry["tp"] if entry else None),
        ("Max Entry", entry["max_entry"] if entry else None),
        ("距 Max Entry %", entry["dist_pct"] if entry else None),
        ("實際盈虧比", entry["ratio"] if entry else None),
        ("─── PER × EPS 換算 ───", ""),
        ("近 4Q EPS", per_summary["eps_4q"] if per_summary else None),
        ("現價對應 PER", per_summary["cur_pe"] if per_summary else None),
        ("保守 18x 合理價", per_summary["conservative_18x"] if per_summary else None),
        ("中間 22x 合理價", per_summary["fair_22x"] if per_summary else None),
        ("─── 事件雷達 ───", ""),
        ("外資 5d 淨(千股)", round(trend["f5"]/1000, 0) if trend["f5"] is not None else None),
        ("外資 20d 淨(千股)", round(trend["f20"]/1000, 0) if trend["f20"] is not None else None),
        ("外資 5d↔20d", trend["flip"] or "同向"),
        ("外資持股 5d 變化 (pp)", trend["sh_change"]),
        ("─── 籌碼 90d ───", ""),
        ("外資淨(千股)", round((f_net or 0)/1000, 0) if f_net else None),
        ("投信淨(千股)", round((t_net or 0)/1000, 0) if t_net else None),
        ("自營淨(千股)", round((d_net or 0)/1000, 0) if d_net else None),
        ("三大合計(千股)", round(((f_net or 0)+(t_net or 0)+(d_net or 0))/1000, 0)),
        ("外資持股%", foreign_pct),
        ("─── 散戶 ───", ""),
        ("融資餘額(張)", round((margin_balance or 0)/1000, 0) if margin_balance else None),
        ("融券餘額(張)", round((short_balance or 0)/1000, 0) if short_balance else None),
        ("─── 警戒 ───", ""),
        ("1y 處置", disp_count),
        ("1y 暫停", suspend_count),
    ]
    sheets["概覽"] = pd.DataFrame(ov, columns=["項目","值"])

    # === 月營收 10Y ===
    if not mrev.empty and "revenue" in mrev.columns:
        m = mrev.sort_values("date").copy().reset_index(drop=True)
        m["營收(億)"] = (m["revenue"] / 1e8).round(2)
        # 手算 YoY (前 12 月同月比較) + MoM
        m["YoY%"] = (m["revenue"] / m["revenue"].shift(12) - 1) * 100
        m["YoY%"] = m["YoY%"].round(1)
        m["MoM%"] = (m["revenue"] / m["revenue"].shift(1) - 1) * 100
        m["MoM%"] = m["MoM%"].round(1)
        sheets["月營收10Y"] = m[["date","營收(億)","YoY%","MoM%"]].sort_values("date", ascending=False)

    # === 損益季 10Y (pivot) ===
    inc = _inc
    if not inc.empty and "type" in inc.columns:
        keep = ["Revenue","CostOfGoodsSold","GrossProfit","OperatingExpenses",
                "OperatingIncome","IncomeAfterTaxes","EPS",
                "ResearchAndDevelopmentExpenses","IncomeAttributableToOwnersOfParent",
                "TotalNonoperatingIncomeAndExpense","IncomeBeforeIncomeTax"]
        i2 = inc[inc["type"].isin(keep)]
        if not i2.empty:
            piv = i2.pivot_table(index="date", columns="type", values="value", aggfunc="last")
            sheets["損益季10Y"] = piv.sort_index(ascending=False).reset_index()

    # === PER 換算 ===
    if per_df is not None:
        sheets["PER換算"] = per_df

    # === 資產負債季 ===
    bs = data.get("bs", pd.DataFrame())
    if not bs.empty and "type" in bs.columns:
        keep = ["CashAndCashEquivalents","Inventories","AccountsReceivableNet",
                "TotalCurrentAssets","PropertyPlantAndEquipment","TotalAssets",
                "AccountsPayable","ShortTermBorrowings","TotalCurrentLiabilities",
                "LongTermLiabilities","Liabilities","TotalLiabilities","Equity",
                "EquityAttributableToOwnersOfParent"]
        b2 = bs[bs["type"].isin(keep)]
        if not b2.empty:
            piv = b2.pivot_table(index="date", columns="type", values="value", aggfunc="last")
            sheets["資產負債季10Y"] = piv.sort_index(ascending=False).reset_index()

    # === 現金流季 ===
    cf = data.get("cf", pd.DataFrame())
    if not cf.empty and "type" in cf.columns:
        keep = ["CashFlowsFromOperatingActivities","CashFlowsFromInvestingActivities",
                "CashFlowsFromFinancingActivities","PropertyPlantAndEquipment",
                "AcquisitionOfPropertyPlantAndEquipment"]
        c2 = cf[cf["type"].isin(keep)]
        if not c2.empty:
            piv = c2.pivot_table(index="date", columns="type", values="value", aggfunc="last")
            sheets["現金流季10Y"] = piv.sort_index(ascending=False).reset_index()

    # === 年營收成長 + CAGR ===
    if not mrev.empty:
        m = mrev.copy(); m["year"] = m["date"].astype(str).str[:4]
        yearly = m.groupby("year").agg(rev=("revenue","sum"), months=("date","count"))
        yearly = yearly[yearly["months"] >= 12].copy()
        yearly["營收(億)"] = (yearly["rev"] / 1e8).round(2)
        yearly["YoY%"] = (yearly["營收(億)"].pct_change() * 100).round(1)
        sheets["年營收成長"] = yearly.reset_index()[["year","營收(億)","YoY%"]]
        # CAGR
        cagr_rows = []
        if len(yearly) >= 4:
            latest_y = yearly.index[-1]; latest_v = yearly.iloc[-1]["營收(億)"]
            for label, n in [("3y", 3), ("5y", 5), ("10y", 10)]:
                if len(yearly) > n:
                    start_v = yearly.iloc[-n-1]["營收(億)"]
                    start_y = yearly.index[-n-1]
                    if start_v > 0 and latest_v > 0:
                        c = round(((latest_v/start_v)**(1/n) - 1) * 100, 1)
                        cagr_rows.append((label, f"{start_y}→{latest_y}", c))
        if cagr_rows:
            sheets["營收CAGR"] = pd.DataFrame(cagr_rows, columns=["期間","區間","CAGR%"])

    # === PER/PBR 歷史 ===
    if not per.empty:
        keep_cols = [c for c in ["date","PER","PBR","dividend_yield"] if c in per.columns]
        sheets["PER_PBR"] = per[keep_cols].sort_values("date", ascending=False)

    # === 市值歷史 ===
    if not mv.empty and "market_value" in mv.columns:
        mv2 = mv.copy()
        mv2["市值(億)"] = (mv2["market_value"] / 1e8).round(1)
        sheets["市值歷史"] = mv2.sort_values("date", ascending=False)[["date","市值(億)"]]

    # === 三大法人 90d ===
    if not iw.empty:
        keep = [c for c in ["date","Foreign_Investor_buy","Foreign_Investor_sell",
                             "Investment_Trust_buy","Investment_Trust_sell",
                             "Dealer_self_buy","Dealer_self_sell",
                             "Dealer_Hedging_buy","Dealer_Hedging_sell"] if c in iw.columns]
        if keep: sheets["三大法人90d"] = iw[keep].sort_values("date", ascending=False)

    # === 外資持股 ===
    if not sh.empty:
        keep = [c for c in ["date","ForeignInvestmentSharesRatio","ForeignInvestmentRemainRatio",
                             "ForeignInvestmentUpperLimitRatio","NumberOfSharesIssued"] if c in sh.columns]
        if keep: sheets["外資持股"] = sh[keep].sort_values("date", ascending=False)

    # === 八大行庫 ===
    gb = data.get("gov_bank", pd.DataFrame())
    if not gb.empty:
        if "stock_id" in gb.columns:
            gb = gb[gb["stock_id"].astype(str) == str(sid)]
        if not gb.empty:
            sheets["八大行庫"] = gb.sort_values("date", ascending=False)

    # === 融資融券 ===
    if not mg.empty:
        keep = [c for c in ["date","MarginPurchaseBuy","MarginPurchaseSell",
                             "MarginPurchaseTodayBalance","ShortSaleBuy",
                             "ShortSaleSell","ShortSaleTodayBalance"] if c in mg.columns]
        if keep: sheets["融資融券"] = mg[keep].sort_values("date", ascending=False)

    # === 借券 ===
    lending = data.get("lending", pd.DataFrame())
    if not lending.empty:
        sheets["借券90d"] = lending.sort_values("date", ascending=False)

    # === 股權分級 ===
    sp = data.get("shares_per", pd.DataFrame())
    if not sp.empty:
        sheets["股權分級"] = sp.sort_values("date", ascending=False)

    # === 股利歷史 ===
    div = data.get("dividend", pd.DataFrame())
    if not div.empty:
        sheets["股利歷史"] = div.sort_values("date", ascending=False)

    # === 除權除息 ===
    dr = data.get("div_result", pd.DataFrame())
    if not dr.empty:
        sheets["除權除息結果"] = dr.sort_values("date", ascending=False)

    # === 拆股/減資 ===
    split = data.get("split", pd.DataFrame())
    if not split.empty:
        sheets["拆股"] = split.sort_values("date", ascending=False)
    capred = data.get("capred", pd.DataFrame())
    if not capred.empty:
        sheets["減資"] = capred.sort_values("date", ascending=False)

    # === 警戒 ===
    disp = data.get("disp", pd.DataFrame())
    if not disp.empty:
        sheets["處置歷史"] = disp.sort_values("date", ascending=False)
    suspend = data.get("suspend", pd.DataFrame())
    if not suspend.empty:
        sheets["暫停歷史"] = suspend.sort_values("date", ascending=False)
    mss = data.get("ms_suspend", pd.DataFrame())
    if not mss.empty:
        sheets["融券暫停"] = mss.sort_values("date", ascending=False)

    # === 同業 (從產業鏈) ===
    if not ind_dataset.empty and "stock_id" in ind_dataset.columns:
        my_ind = ind_dataset[ind_dataset["stock_id"].astype(str) == str(sid)]
        if len(my_ind) and "industry_category" in my_ind.columns:
            my_cats = set(my_ind["industry_category"].astype(str))
            peers = ind_dataset[ind_dataset["industry_category"].astype(str).isin(my_cats)]
            peers = peers[peers["stock_id"].astype(str) != str(sid)]
            if not peers.empty:
                sheets["同業"] = peers.drop_duplicates(subset="stock_id").head(30)

    # === 新聞 ===
    news = data.get("news", pd.DataFrame())
    if not news.empty:
        keep = [c for c in ["date","title","description","link","source"] if c in news.columns]
        if keep: sheets["新聞"] = news[keep].sort_values("date", ascending=False).head(30)

    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with pd.ExcelWriter(dst, engine="openpyxl") as xw:
        for sname, df in sheets.items():
            df.to_excel(xw, sheet_name=sname[:31], index=False)
    return sheets


def _calc_entry_tw(price, yr_low, yr_high):
    """3:1 計算, 返回 dict 或 None
    SL = 52w 低 × 1.03, TP = 52w 高 × 1.05
    """
    if not (price and yr_low and yr_high and yr_high > yr_low):
        return None
    sl = yr_low * 1.03
    tp = yr_high * 1.05
    if tp <= sl:
        return None
    max_entry = (tp + 3 * sl) / 4
    if price <= sl: verdict = "🚨 已破止損(SL 失守)"
    elif price <= max_entry: verdict = "🟢 進場區(≥3:1)"
    elif price <= max_entry * 1.05: verdict = "🟡 接近門檻(<5%)"
    elif price <= max_entry * 1.15: verdict = "🟠 稍高(5-15%)"
    else: verdict = "🔴 追高風險(>15%)"
    ratio = round((tp - price) / (price - sl), 2) if price > sl else None
    dist_pct = round((price / max_entry - 1) * 100, 1)
    return {"sl": round(sl, 2), "tp": round(tp, 2), "max_entry": round(max_entry, 2),
            "verdict": verdict, "ratio": ratio, "dist_pct": dist_pct}


def _foreign_trend_tw(iw, sh):
    """外資 5d/20d 動向 + 持股 % 5d 變化
    返回 dict: {f5, f20, flip, sh_change}
    """
    out = {"f5": None, "f20": None, "flip": None, "sh_change": None}
    if not iw.empty and "Foreign_Investor_buy" in iw.columns and "Foreign_Investor_sell" in iw.columns:
        iw_s = iw.sort_values("date", ascending=False)
        def net(df, n):
            head = df.head(n)
            return int((head["Foreign_Investor_buy"].fillna(0)
                        - head["Foreign_Investor_sell"].fillna(0)).sum())
        if len(iw_s) >= 5:  out["f5"]  = net(iw_s, 5)
        if len(iw_s) >= 20: out["f20"] = net(iw_s, 20)
        if out["f5"] is not None and out["f20"] is not None:
            if out["f5"] < 0 < out["f20"]:
                out["flip"] = "🚨 5d 轉賣 (背離 20d)"
            elif out["f5"] > 0 > out["f20"]:
                out["flip"] = "🟢 5d 轉買 (背離 20d)"
    if not sh.empty and "ForeignInvestmentSharesRatio" in sh.columns:
        sh_s = sh.sort_values("date", ascending=False)
        if len(sh_s) >= 5:
            latest = float(sh_s.iloc[0]["ForeignInvestmentSharesRatio"])
            past = float(sh_s.iloc[4]["ForeignInvestmentSharesRatio"])
            out["sh_change"] = round(latest - past, 2)
    return out


def _event_radar_tw_md(entry, trend):
    if not entry and not any(v is not None for v in trend.values()):
        return ""
    lines = ["\n## 🚨 事件雷達 (短期籌碼)\n",
             "| 訊號 | 值 | 判讀 |", "|---|---:|---|"]
    if entry:
        lines.append(f"| **3:1 判讀** | 距 Max Entry {entry['dist_pct']:+.1f}% | **{entry['verdict']}** |")
    if trend["f5"] is not None:
        lines.append(f"| 外資 5d 淨 | {trend['f5']/1000:+,.0f} 張 | {'🟢 買' if trend['f5']>0 else '🔴 賣'} |")
    if trend["f20"] is not None:
        lines.append(f"| 外資 20d 淨 | {trend['f20']/1000:+,.0f} 張 | {'🟢 買' if trend['f20']>0 else '🔴 賣'} |")
    if trend["flip"]:
        lines.append(f"| 5d ↔ 20d | {trend['flip']} | ⚠️ |")
    if trend["sh_change"] is not None:
        emoji = '🟢' if trend['sh_change'] > 0.1 else '🔴' if trend['sh_change'] < -0.1 else '⚪'
        lines.append(f"| 外資持股 5d 變化 | {trend['sh_change']:+.2f} pp | {emoji} |")
    return "\n".join(lines) + "\n"


def _summary_action_tw(entry, trend):
    if not entry:
        return ""
    md = f"""
---

## 🎯 綜合判讀 / 事件行動

### 進場策略 (3:1 派)
- **限價買**: {entry['max_entry']}
- **止損**: {entry['sl']} (跌破止損, 重新評估)
- **目標**: {entry['tp']}
- **當前判讀**: {entry['verdict']}
"""
    warnings = []
    if trend["flip"]:
        warnings.append(trend["flip"])
    if trend["sh_change"] is not None and abs(trend["sh_change"]) > 0.5:
        warnings.append(f"外資持股 5d 變化 {trend['sh_change']:+.2f} pp (幅度顯著)")
    if warnings:
        md += "\n### 🚨 短期警訊\n" + "\n".join(f"- {w}" for w in warnings) + "\n"
    md += """
### 撤退訊號 (任一觸發即檢討)
1. 月營收 YoY 連 2 月轉負
2. EPS 連 2 季 QoQ 下滑
3. 外資持股 % 跌破近 3M 低點
4. 股價跌破 SL

⚠️ 3:1 SL/TP 用 52w 高低粗略估算 (成長股會失真, 高波動股會太寬鬆), 實戰請自行核對技術支撐 + 基本面目標
"""
    return md


def _per_band_tw(price, inc):
    """PER × 近 4Q EPS 換算價. 返回 (md_str, df) 或 (None, None)"""
    if inc.empty or "type" not in inc.columns:
        return None, None
    eps_rows = inc[inc["type"] == "EPS"].sort_values("date", ascending=False)
    if len(eps_rows) < 4 or not price:
        return None, None
    eps_4q = float(eps_rows.head(4)["value"].sum())
    if eps_4q <= 0:
        return None, None
    bands = [10, 14, 18, 22, 26, 30]
    rows = []
    for pe in bands:
        tgt = round(eps_4q * pe, 2)
        diff = round((price / tgt - 1) * 100, 1)
        if diff < -20:   v = "🟢 大幅低估"
        elif diff < -5:  v = "🟢 低估"
        elif diff < 5:   v = "🟡 合理"
        elif diff < 20:  v = "🟠 高估"
        else:            v = "🔴 大幅高估"
        rows.append({"PER 倍": f"{pe}x", "對應價": tgt, "距現價 %": diff, "判讀": v})
    df = pd.DataFrame(rows)
    cur_pe = round(price / eps_4q, 1)
    md = f"""
## 📊 PER 換算價 (基於近 4Q EPS)

- 近 4Q EPS 合計: **{eps_4q:.2f}** 元
- 現價 {price} → 對應 **PER {cur_pe}x**

| PER | 對應價 | 距現價 % | 判讀 |
|---:|---:|---:|---|
"""
    for r in rows:
        s = '+' if r['距現價 %'] > 0 else ''
        md += f"| {r['PER 倍']} | {r['對應價']} | {s}{r['距現價 %']}% | {r['判讀']} |\n"
    md += """
⚠️ 假設 EPS 維持. 若 EPS 加速成長 → 高倍數也合理; 若衰退 → 低倍數才安全
💡 3:1 (技術/籌碼) vs PER (基本面/估值) 兩個視角常會給出不同答案 — 用來 cross check
"""
    return md, df


def _entry_section_tw(price, yr_low, yr_high, num):
    """3:1 盈虧比入場計算段落 (April1Stock 公式)"""
    e = _calc_entry_tw(price, yr_low, yr_high)
    if not e:
        return ""
    ratio_s = num(e["ratio"]) if e["ratio"] is not None else "—"
    return f"""
## 🎯 3:1 入場計算 (April1Stock 公式)

Max Entry = (TP + 3×SL) / 4 → 現價 ≤ Max Entry 才符合 3:1 盈虧比

| 項目 | 值 | 說明 |
|---|---:|---|
| **判讀** | **{e['verdict']}** | |
| SL 止損 | {num(e['sl'])} | 52w 低 × 1.03 |
| TP 目標 | {num(e['tp'])} | 52w 高 × 1.05 (較粗糙) |
| **Max Entry** | **{num(e['max_entry'])}** | 現價 ≤ 此值才買 |
| 現價 | {num(price)} | 距 Max Entry {'+' if e['dist_pct'] > 0 else ''}{e['dist_pct']}% |
| 實際盈虧比 | {ratio_s} | 目標 ≥ 3.0 |

⚠️ 台股 TP 用 52w 高估算, 較樂觀; 實戰請自行核對基本面目標
"""


def build_md(sid, name, data, dst):
    price_1y = data.get("price_1y", pd.DataFrame())
    per = data.get("per", pd.DataFrame())
    mv = data.get("market_val", pd.DataFrame())
    mrev = data.get("month_rev", pd.DataFrame())
    iw = data.get("inst_wide", pd.DataFrame())
    sh = data.get("shareholding", pd.DataFrame())
    mg = data.get("margin", pd.DataFrame())
    div = data.get("dividend", pd.DataFrame())
    news = data.get("news", pd.DataFrame())
    disp = data.get("disp", pd.DataFrame())
    ind_dataset = data.get("industry", pd.DataFrame())

    def num(v, d=2):
        try: return f"{float(v):,.{d}f}"
        except: return "—"

    latest_price = year_high = year_low = None
    if not price_1y.empty and "close" in price_1y.columns:
        p = price_1y.sort_values("date")
        latest_price = float(p.iloc[-1]["close"])
        year_high = float(p["max"].max()) if "max" in p.columns else float(p["close"].max())
        year_low = float(p["min"].min()) if "min" in p.columns else float(p["close"].min())

    latest_per = latest_pbr = div_y = None
    if not per.empty:
        pers = per.sort_values("date")
        latest_per = pers.iloc[-1].get("PER")
        latest_pbr = pers.iloc[-1].get("PBR")
        div_y = pers.iloc[-1].get("dividend_yield")

    latest_mv = None
    if not mv.empty and "market_value" in mv.columns:
        latest_mv = mv.sort_values("date").iloc[-1]["market_value"]

    rev_yoy = latest_month_rev = None
    if not mrev.empty:
        mrs = mrev.sort_values("date")
        latest_month_rev = mrs.iloc[-1].get("revenue")
        rev_yoy = mrs.iloc[-1].get("revenue_year_growth")

    def s(df, col): return int(df[col].sum()) if col in df.columns else 0
    f_net = t_net = d_net = 0
    if not iw.empty:
        f_net = s(iw,"Foreign_Investor_buy") - s(iw,"Foreign_Investor_sell")
        t_net = s(iw,"Investment_Trust_buy") - s(iw,"Investment_Trust_sell")
        d_net = (s(iw,"Dealer_buy")+s(iw,"Dealer_self_buy")+s(iw,"Dealer_Hedging_buy")
                 - s(iw,"Dealer_sell")-s(iw,"Dealer_self_sell")-s(iw,"Dealer_Hedging_sell"))

    foreign_pct = None
    if not sh.empty and "ForeignInvestmentSharesRatio" in sh.columns:
        foreign_pct = sh.sort_values("date").iloc[-1]["ForeignInvestmentSharesRatio"]

    industry_str = ""
    if not ind_dataset.empty and "stock_id" in ind_dataset.columns:
        my = ind_dataset[ind_dataset["stock_id"].astype(str) == str(sid)]
        if len(my) and "industry_category" in my.columns:
            industry_str = " / ".join(my["industry_category"].astype(str).unique()[:3])

    # CAGR
    cagr_lines = []
    if not mrev.empty:
        m = mrev.copy(); m["year"] = m["date"].astype(str).str[:4]
        yearly = m.groupby("year").agg(rev=("revenue","sum"), months=("date","count"))
        yearly = yearly[yearly["months"] >= 12].sort_index()
        if len(yearly) >= 4:
            latest_v = yearly.iloc[-1]["rev"]
            for label, n in [("3y", 3), ("5y", 5), ("10y", 10)]:
                if len(yearly) > n:
                    start_v = yearly.iloc[-n-1]["rev"]
                    if start_v > 0 and latest_v > 0:
                        c = round(((latest_v/start_v)**(1/n) - 1) * 100, 1)
                        cagr_lines.append((label, c))

    md = f"""# 🔍 {sid} — {name}

**產業:** {industry_str or '—'}

---

## 💰 報價 & 估值

| 項目 | 值 | 項目 | 值 |
|---|---:|---|---:|
| 股價 | **${num(latest_price)}** | 市值 | {round((latest_mv or 0)/1e8, 1) if latest_mv else '—'} 億 |
| 52w 高 | {num(year_high)} | 52w 低 | {num(year_low)} |
| **PER** | **{num(latest_per)}** | PBR | {num(latest_pbr)} |
| 殖利率 | {num(div_y)}% | — | — |
{_event_radar_tw_md(_calc_entry_tw(latest_price, year_low, year_high), _foreign_trend_tw(iw, sh))}
{_entry_section_tw(latest_price, year_low, year_high, num)}
{_per_band_tw(latest_price, data.get('income', pd.DataFrame()))[0] or ''}

## 📈 營收動能

| 項目 | 值 |
|---|---:|
| 最新月營收 | **{round((latest_month_rev or 0)/1e8, 1) if latest_month_rev else '—'} 億** |
| **月營收 YoY** | **{num(rev_yoy)}%** |"""
    for label, c in cagr_lines:
        md += f"\n| 營收 {label} CAGR | **{c}%** |"

    md += f"""

## 🌍 籌碼 (90 天累計)

| 項目 | 千股 |
|---|---:|
| 🌍 外資淨買賣 | **{round(f_net/1000, 0):,.0f}** |
| 📈 投信淨買賣 | **{round(t_net/1000, 0):,.0f}** |
| 🎯 自營淨買賣 | **{round(d_net/1000, 0):,.0f}** |
| **三大合計** | **{round((f_net+t_net+d_net)/1000, 0):,.0f}** |
| 外資持股 % | **{num(foreign_pct)}%** |
"""

    if not mg.empty:
        mgs = mg.sort_values("date")
        mb = mgs.iloc[-1].get("MarginPurchaseTodayBalance")
        sb = mgs.iloc[-1].get("ShortSaleTodayBalance")
        md += f"\n## 💸 散戶\n\n- 融資餘額: **{int((mb or 0)/1000):,} 張**\n- 融券餘額: **{int((sb or 0)/1000):,} 張**\n"

    if not disp.empty:
        md += f"\n## ⚠️ 警戒 (近 1 年)\n\n- 處置次數: **{len(disp)}** 次\n"
        for _, r in disp.sort_values("date", ascending=False).head(3).iterrows():
            md += f"  - {r.get('date','')}: {str(r.get('condition',''))[:50]}\n"

    if not ind_dataset.empty and "stock_id" in ind_dataset.columns:
        my_ind = ind_dataset[ind_dataset["stock_id"].astype(str) == str(sid)]
        if len(my_ind) and "industry_category" in my_ind.columns:
            my_cats = set(my_ind["industry_category"].astype(str))
            peers = ind_dataset[ind_dataset["industry_category"].astype(str).isin(my_cats)]
            peers = peers[peers["stock_id"].astype(str) != str(sid)].drop_duplicates(subset="stock_id")
            if not peers.empty:
                md += "\n## 👥 同業 (產業鏈)\n\n"
                for _, r in peers.head(10).iterrows():
                    md += f"- **{r.get('stock_id','')}** {r.get('stock_name','')} ({r.get('industry_category','')})\n"

    if not div.empty:
        md += "\n## 💵 股利歷史 (近 5 年)\n\n| 年 | 現金 | 股票 |\n|---|---:|---:|"
        d = div.copy()
        d["_year"] = d.get("date", "").astype(str).str[:4]
        d["_cash"] = pd.to_numeric(d.get("CashEarningsDistribution"), errors="coerce").fillna(0)
        d["_stk"] = pd.to_numeric(d.get("StockEarningsDistribution"), errors="coerce").fillna(0)
        by_year = d.groupby("_year").agg(cash=("_cash", "sum"), stk=("_stk", "sum"))
        for y, r in by_year.sort_index(ascending=False).head(5).iterrows():
            md += f"\n| {y} | {r['cash']:.2f} | {r['stk']:.2f} |"

    if not news.empty:
        md += "\n\n## 📰 最新新聞\n\n"
        for _, r in news.sort_values("date", ascending=False).head(5).iterrows():
            title = str(r.get("title",""))[:80]; date = str(r.get("date",""))[:10]
            link = r.get("link","")
            md += f"- [{title}]({link}) — *{date}*\n"

    md += _summary_action_tw(_calc_entry_tw(latest_price, year_low, year_high),
                              _foreign_trend_tw(iw, sh))

    md += f"\n\n---\n\n_資料源: FinMind • 生成 {datetime.now().strftime('%Y-%m-%d %H:%M')}_\n"

    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(md)


def get_name(sid, data):
    info = data.get("info", pd.DataFrame())
    if not info.empty and "stock_id" in info.columns:
        info["stock_id"] = info["stock_id"].astype(str)
        r = info[info["stock_id"] == str(sid)]
        if len(r): return r.iloc[0].get("stock_name", "")
    return ""


def process(sid):
    print(f"\n=== {sid} 深度研究 ===")
    t0 = time.time()
    data = fetch_all(sid)
    ok = sum(1 for v in data.values() if not v.empty)
    print(f"  抓 {ok}/{len(data)} dataset ({time.time()-t0:.1f}s)")
    name = get_name(sid, data)
    print(f"  名稱: {name}")
    xlsx_dst = os.path.join(OUT_DIR, f"台股深度_{sid}.xlsx")
    md_dst = os.path.join(OUT_DIR, f"台股深度_{sid}.md")
    sheets = build_xlsx(sid, name, data, xlsx_dst)
    build_md(sid, name, data, md_dst)
    print(f"  → {xlsx_dst}  ({len(sheets)} 分頁)")
    print(f"  → {md_dst}")


def main():
    if not TOKEN: print("⚠️ 需 FINMIND_TOKEN"); sys.exit(1)
    args = [a for a in sys.argv[1:] if a.strip()]
    if not args:
        env = os.environ.get("TICKER", "").strip()
        if env: args = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
    if not args:
        print("⚠️ 未指定代號: python tw_stock_deepdive.py 2330"); sys.exit(1)
    for sid in args: process(sid)


if __name__ == "__main__":
    main()
