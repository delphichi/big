# -*- coding: utf-8 -*-
"""
美股單檔深度研究 us_stock_deepdive.py
=======================================================================
對單一美股跑 ~40 個 FMP 端點, 整合成 Bloomberg-style 完整研究報告

抓內容:
  1. 公司概況 (profile, sector, CEO, 員工, 市值)
  2. 報價 & 漲跌 (1D/5D/1M/3M/6M/1Y/3Y/5Y/10Y)
  3. 三大報表 10 年 (annual)
  4. 成長率 (YoY + 3Y/5Y/10Y CAGR)
  5. 財務比率 10 年 (毛利率/淨利率/ROE/ROA/ROIC 等)
  6. 估值 (P/E / PEG / EV/EBITDA / DCF / Levered DCF / Owner Earnings)
  7. 財務評分 (Altman Z / Piotroski)
  8. 分析師預估 (EPS/營收未來 3 年 + 目標價共識 + 評級)
  9. 產品/地理營收結構 (歷年)
  10. 內部人交易 + 國會交易
  11. 5% 大戶持股異動
  12. 員工數變化 / 高管
  13. 股利 歷史
  14. 技術面 (SMA50/200, RSI14)
  15. 同業對比 (stock-peers)
  16. 個股新聞 (10 則)
  17. SEC 8-K 重大事件

跑法:
  FMP_API_KEY=xxx python us_stock_deepdive.py NVDA
  FMP_API_KEY=xxx python us_stock_deepdive.py NVDA META GOOG  # 多檔
  TICKER=NVDA python us_stock_deepdive.py                     # env 也可

輸出 (每檔):
  data/deepdive/深度研究_{TICKER}.xlsx  (~20 分頁完整資料)
  data/deepdive/深度研究_{TICKER}.md    (Markdown 精華, 可貼 email/GitHub)
"""
import os, sys, time, requests
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
OUT_DIR = os.environ.get("OUT_DIR", "data/deepdive")


def get(endpoint, **params):
    if not KEY: return None
    params["apikey"] = KEY
    for i in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=15)
            if r.status_code == 429: time.sleep(2*(i+1)); continue
            if r.status_code != 200: return None
            return r.json()
        except Exception: time.sleep(1)
    return None


def first(d):
    if d is None: return None
    if isinstance(d, list): return d[0] if d else None
    return d


def to_pct(v):
    if v is None: return None
    try: return round(float(v) * 100, 2)
    except: return None


def fetch_all(sym):
    """所有 endpoint 平行抓"""
    d90 = (datetime.now()-timedelta(days=90)).strftime("%Y-%m-%d")
    d3y = (datetime.now()-timedelta(days=365*3)).strftime("%Y-%m-%d")
    tasks = [
        ("profile",         ("profile", {"symbol": sym})),
        ("quote",           ("quote", {"symbol": sym})),
        ("price_change",    ("stock-price-change", {"symbol": sym})),
        ("mcap_hist",       ("historical-market-capitalization", {"symbol": sym, "from": d3y, "limit": 40})),
        ("shares_float",    ("shares-float", {"symbol": sym})),
        ("key_executives",  ("key-executives", {"symbol": sym})),
        ("emp_hist",        ("historical-employee-count", {"symbol": sym, "limit": 10})),
        # 財報 10Y
        ("income",  ("income-statement", {"symbol": sym, "period": "annual", "limit": 10})),
        ("bs",      ("balance-sheet-statement", {"symbol": sym, "period": "annual", "limit": 10})),
        ("cf",      ("cash-flow-statement", {"symbol": sym, "period": "annual", "limit": 10})),
        # 成長 / 比率
        ("inc_growth",  ("income-statement-growth", {"symbol": sym, "period": "annual", "limit": 5})),
        ("fin_growth",  ("financial-growth", {"symbol": sym, "period": "annual", "limit": 3})),
        ("ratios",      ("ratios", {"symbol": sym, "period": "annual", "limit": 10})),
        ("ratios_ttm",  ("ratios-ttm", {"symbol": sym})),
        ("key_metrics", ("key-metrics", {"symbol": sym, "period": "annual", "limit": 10})),
        ("km_ttm",      ("key-metrics-ttm", {"symbol": sym})),
        ("ev",          ("enterprise-values", {"symbol": sym, "period": "annual", "limit": 5})),
        # 估值
        ("dcf",     ("discounted-cash-flow", {"symbol": sym})),
        ("dcf_lev", ("levered-discounted-cash-flow", {"symbol": sym})),
        ("owner_e", ("owner-earnings", {"symbol": sym, "limit": 5})),
        ("scores",  ("financial-scores", {"symbol": sym})),
        # 分析師
        ("estimates",       ("analyst-estimates", {"symbol": sym, "period": "annual", "limit": 3})),
        ("target_summary",  ("price-target-summary", {"symbol": sym})),
        ("target_consensus",("price-target-consensus", {"symbol": sym})),
        ("grades_consensus",("grades-consensus", {"symbol": sym})),
        ("grades_hist",     ("grades-historical", {"symbol": sym, "limit": 10})),
        ("ratings_snap",    ("ratings-snapshot", {"symbol": sym})),
        # 事件
        ("earnings",    ("earnings", {"symbol": sym, "limit": 6})),
        # 營收結構
        ("prod_seg",    ("revenue-product-segmentation", {"symbol": sym, "period": "annual", "limit": 5})),
        ("geo_seg",     ("revenue-geographic-segmentation", {"symbol": sym, "period": "annual", "limit": 5})),
        # 內部人 / 國會 / 大戶
        ("insider_stats", ("insider-trading/statistics", {"symbol": sym})),
        ("insider_search",("insider-trading/search", {"symbol": sym, "limit": 15})),
        ("beneficial",    ("acquisition-of-beneficial-ownership", {"symbol": sym})),
        ("senate",        ("senate-trades", {"symbol": sym})),
        ("house",         ("house-trades", {"symbol": sym})),
        # 同業 / 股利
        ("peers",       ("stock-peers", {"symbol": sym})),
        ("dividends",   ("dividends", {"symbol": sym, "limit": 10})),
        ("splits",      ("splits", {"symbol": sym, "limit": 5})),
        # 技術面
        ("sma50",  ("technical-indicators/sma", {"symbol": sym, "periodLength": 50, "timeframe": "1day", "from": d90})),
        ("sma200", ("technical-indicators/sma", {"symbol": sym, "periodLength": 200, "timeframe": "1day", "from": d90})),
        ("rsi",    ("technical-indicators/rsi", {"symbol": sym, "periodLength": 14, "timeframe": "1day", "from": d90})),
        # 新聞 / SEC
        ("news",    ("news/stock", {"symbols": sym, "limit": 10})),
        ("sec_8k",  ("sec-filings-8k", {"symbol": sym, "limit": 5})),
    ]

    out = {}
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = {ex.submit(get, ep, **params): key for key, (ep, params) in tasks}
        for fut in as_completed(futs):
            key = futs[fut]
            try: out[key] = fut.result()
            except Exception: out[key] = None
    return out


def cumul_to_cagr(c, n):
    if c is None: return None
    try:
        if (1 + float(c)) <= 0: return None
        return round(((1 + float(c)) ** (1/n) - 1) * 100, 2)
    except: return None


def build_xlsx(sym, data, dst):
    """多分頁 xlsx"""
    sheets = {}

    prof  = first(data.get("profile")) or {}
    quote = first(data.get("quote")) or {}
    pc    = first(data.get("price_change")) or {}
    scores= first(data.get("scores")) or {}
    tgt   = first(data.get("target_consensus")) or {}
    grades= first(data.get("grades_consensus")) or {}
    dcf   = first(data.get("dcf")) or {}
    dcf_l = first(data.get("dcf_lev")) or {}
    kmttm = first(data.get("km_ttm")) or {}
    rttm  = first(data.get("ratios_ttm")) or {}
    fg    = first(data.get("fin_growth")) or {}
    price = quote.get("price")

    # === 概覽 ===
    ov = [
        ("代號", sym),
        ("名稱", prof.get("companyName")),
        ("產業", prof.get("industry")),
        ("Sector", prof.get("sector")),
        ("國家", prof.get("country")),
        ("CEO", prof.get("ceo")),
        ("員工數", prof.get("fullTimeEmployees")),
        ("IPO", prof.get("ipoDate")),
        ("Beta", prof.get("beta")),
        ("網站", prof.get("website")),
        ("─── 報價 ───", ""),
        ("當前股價", price),
        ("52w 高", quote.get("yearHigh")),
        ("52w 低", quote.get("yearLow")),
        ("市值(億)", round((quote.get("marketCap") or 0)/1e8, 1) if quote.get("marketCap") else None),
        ("MA50", quote.get("priceAvg50")),
        ("MA200", quote.get("priceAvg200")),
        ("─── 漲跌 ───", ""),
        ("1D%", pc.get("1D")), ("1M%", pc.get("1M")), ("3M%", pc.get("3M")),
        ("YTD%", pc.get("ytd")), ("1Y%", pc.get("1Y")),
        ("3Y%", pc.get("3Y")), ("5Y%", pc.get("5Y")), ("10Y%", pc.get("10Y")),
        ("─── 估值 ───", ""),
        ("P/E TTM", rttm.get("priceToEarningsRatioTTM")),
        ("Fwd P/E", rttm.get("forwardPriceToEarningsGrowthRatioTTM")),
        ("PEG TTM", rttm.get("priceToEarningsGrowthRatioTTM")),
        ("P/B TTM", rttm.get("priceToBookRatioTTM")),
        ("P/S TTM", rttm.get("priceToSalesRatioTTM")),
        ("EV/EBITDA TTM", kmttm.get("evToEBITDATTM")),
        ("EV/FCF TTM", kmttm.get("evToFreeCashFlowTTM")),
        ("DCF", dcf.get("dcf")),
        ("Levered DCF", dcf_l.get("dcf")),
        ("DCF 差 %", round((dcf.get("dcf")/price-1)*100,1) if price and dcf.get("dcf") else None),
        ("Graham #", kmttm.get("grahamNumberTTM")),
        ("FCF Yield %", to_pct(kmttm.get("freeCashFlowYieldTTM"))),
        ("Earnings Yield %", to_pct(kmttm.get("earningsYieldTTM"))),
        ("─── 品質 ───", ""),
        ("ROE TTM %", to_pct(rttm.get("returnOnEquityTTM"))),
        ("ROA TTM %", to_pct(rttm.get("returnOnAssetsTTM"))),
        ("ROIC TTM %", to_pct(kmttm.get("returnOnInvestedCapitalTTM"))),
        ("毛利率 TTM %", to_pct(rttm.get("grossProfitMarginTTM"))),
        ("營益率 TTM %", to_pct(rttm.get("operatingProfitMarginTTM"))),
        ("淨利率 TTM %", to_pct(rttm.get("netProfitMarginTTM"))),
        ("Altman Z", scores.get("altmanZScore")),
        ("Piotroski F", scores.get("piotroskiScore")),
        ("FMP 評等", (first(data.get("ratings_snap")) or {}).get("rating")),
        ("─── 分析師 ───", ""),
        ("目標價中位", tgt.get("targetMedian")),
        ("目標價高", tgt.get("targetHigh")),
        ("目標價低", tgt.get("targetLow")),
        ("目標價 vs 現價 %", round((tgt.get("targetMedian")/price-1)*100,1) if price and tgt.get("targetMedian") else None),
        ("Strong Buy", grades.get("strongBuy")),
        ("Buy", grades.get("buy")),
        ("Hold", grades.get("hold")),
        ("Sell", grades.get("sell")),
        ("華爾街共識", grades.get("consensus")),
    ]

    # 3:1 入場 + 事件雷達 + PER 換算
    entry = _calc_entry_us(price, quote.get("priceAvg200"), quote.get("yearLow"), tgt.get("targetLow"))
    sig = _us_signals(data)
    per_md, per_df = _per_band_us(price, data)
    per_eps = per_cur_pe = per_fair_20x = per_conservative_15x = None
    if per_df is not None and not per_df.empty:
        b15 = per_df[per_df["PER 倍"] == "15x"]
        b20 = per_df[per_df["PER 倍"] == "20x"]
        per_conservative_15x = float(b15.iloc[0]["對應價"]) if len(b15) else None
        per_fair_20x = float(b20.iloc[0]["對應價"]) if len(b20) else None
        per_eps = round(per_fair_20x / 20, 2) if per_fair_20x else None
        per_cur_pe = round(price / per_eps, 1) if per_eps and price else None
    ov.extend([
        ("─── 3:1 入場 ───", ""),
        ("3:1 判讀", entry["verdict"] if entry else None),
        ("SL 止損", entry["sl"] if entry else None),
        ("TP 目標", entry["tp"] if entry else None),
        ("Max Entry", entry["max_entry"] if entry else None),
        ("距 Max Entry %", entry["dist_pct"] if entry else None),
        ("實際盈虧比", entry["ratio"] if entry else None),
        ("─── 事件雷達 ───", ""),
        ("內部人 4Q 買賣比", sig["insider_ratio"]),
        ("內部人訊號", sig["insider_signal"]),
        ("國會 90d 買筆", sig["cong_buy"]),
        ("國會 90d 賣筆", sig["cong_sell"]),
        ("國會訊號", sig["cong_signal"]),
        ("─── PER × EPS 換算 ───", ""),
        ("TTM EPS", per_eps),
        ("現價對應 PER", per_cur_pe),
        ("保守 15x 合理價", per_conservative_15x),
        ("中間 20x 合理價", per_fair_20x),
    ])
    sheets["概覽"] = pd.DataFrame(ov, columns=["項目","值"])
    if per_df is not None:
        sheets["PER換算"] = per_df

    # === 損益 10Y ===
    inc = data.get("income") or []
    if inc:
        rows = [{
            "年": (r.get("date") or "")[:4],
            "營收": r.get("revenue"), "毛利": r.get("grossProfit"),
            "營業利益": r.get("operatingIncome"), "EBITDA": r.get("ebitda"), "EBIT": r.get("ebit"),
            "淨利": r.get("netIncome"), "EPS": r.get("eps"), "EPS稀釋": r.get("epsDiluted"),
            "股數(在外)": r.get("weightedAverageShsOut"),
            "R&D": r.get("researchAndDevelopmentExpenses"),
            "利息費用": r.get("interestExpense"), "稅": r.get("incomeTaxExpense"),
        } for r in inc]
        sheets["損益10Y"] = pd.DataFrame(rows).sort_values("年", ascending=False)

    # === BS 10Y ===
    bs = data.get("bs") or []
    if bs:
        rows = [{
            "年": (r.get("date") or "")[:4],
            "現金": r.get("cashAndCashEquivalents"), "短期投資": r.get("shortTermInvestments"),
            "應收": r.get("netReceivables"), "存貨": r.get("inventory"),
            "流動資產": r.get("totalCurrentAssets"), "PPE": r.get("propertyPlantEquipmentNet"),
            "商譽": r.get("goodwill"), "總資產": r.get("totalAssets"),
            "應付": r.get("accountPayables"), "短期債": r.get("shortTermDebt"),
            "流動負債": r.get("totalCurrentLiabilities"), "長期債": r.get("longTermDebt"),
            "總負債": r.get("totalLiabilities"), "股東權益": r.get("totalStockholdersEquity"),
        } for r in bs]
        sheets["資產負債10Y"] = pd.DataFrame(rows).sort_values("年", ascending=False)

    # === CF 10Y ===
    cf = data.get("cf") or []
    if cf:
        rows = [{
            "年": (r.get("date") or "")[:4],
            "淨利": r.get("netIncome"), "折舊": r.get("depreciationAndAmortization"),
            "SBC": r.get("stockBasedCompensation"),
            "OCF": r.get("netCashProvidedByOperatingActivities"),
            "CapEx": r.get("investmentsInPropertyPlantAndEquipment"),
            "投資 CF": r.get("netCashProvidedByInvestingActivities"),
            "還債": r.get("debtRepayment"), "發股": r.get("commonStockIssued"),
            "買回": r.get("commonStockRepurchased"), "股利": r.get("dividendsPaid"),
            "融資 CF": r.get("netCashProvidedByFinancingActivities"),
            "FCF": r.get("freeCashFlow"),
        } for r in cf]
        sheets["現金流10Y"] = pd.DataFrame(rows).sort_values("年", ascending=False)

    # === 成長率 ===
    g_rows = [("─── 1Y YoY ───", "")]
    if fg:
        g_rows += [
            ("營收 YoY %", to_pct(fg.get("revenueGrowth"))),
            ("毛利 YoY %", to_pct(fg.get("grossProfitGrowth"))),
            ("營益 YoY %", to_pct(fg.get("operatingIncomeGrowth"))),
            ("EBIT YoY %", to_pct(fg.get("ebitgrowth"))),
            ("淨利 YoY %", to_pct(fg.get("netIncomeGrowth"))),
            ("EPS YoY %", to_pct(fg.get("epsgrowth"))),
            ("EPS稀釋 YoY %", to_pct(fg.get("epsdilutedGrowth"))),
            ("R&D YoY %", to_pct(fg.get("rdexpenseGrowth"))),
            ("OCF YoY %", to_pct(fg.get("operatingCashFlowGrowth"))),
            ("FCF YoY %", to_pct(fg.get("freeCashFlowGrowth"))),
            ("存貨 YoY %", to_pct(fg.get("inventoryGrowth"))),
            ("應收 YoY %", to_pct(fg.get("receivablesGrowth"))),
            ("資產 YoY %", to_pct(fg.get("assetGrowth"))),
            ("債 YoY %", to_pct(fg.get("debtGrowth"))),
            ("股利/股 YoY %", to_pct(fg.get("dividendsPerShareGrowth"))),
            ("股數 YoY %", to_pct(fg.get("weightedAverageSharesGrowth"))),
        ]
        for label, prefix, n in [("─── 3Y CAGR (per share) ───","threeY",3),
                                   ("─── 5Y CAGR (per share) ───","fiveY",5),
                                   ("─── 10Y CAGR (per share) ───","tenY",10)]:
            g_rows.append((label, ""))
            g_rows += [
                (f"營收/股 {n}Y CAGR %", cumul_to_cagr(fg.get(f"{prefix}RevenueGrowthPerShare"), n)),
                (f"淨利/股 {n}Y CAGR %", cumul_to_cagr(fg.get(f"{prefix}NetIncomeGrowthPerShare"), n)),
                (f"OCF/股 {n}Y CAGR %", cumul_to_cagr(fg.get(f"{prefix}OperatingCFGrowthPerShare"), n)),
                (f"股利/股 {n}Y CAGR %", cumul_to_cagr(fg.get(f"{prefix}DividendperShareGrowthPerShare"), n)),
                (f"權益/股 {n}Y CAGR %", cumul_to_cagr(fg.get(f"{prefix}ShareholdersEquityGrowthPerShare"), n)),
            ]
    sheets["成長率"] = pd.DataFrame(g_rows, columns=["項目","值"])

    # === 比率 10Y ===
    ratios = data.get("ratios") or []
    if ratios:
        rows = [{
            "年": (r.get("date") or "")[:4],
            "毛利率%": to_pct(r.get("grossProfitMargin")),
            "營益率%": to_pct(r.get("operatingProfitMargin")),
            "淨利率%": to_pct(r.get("netProfitMargin")),
            "EBITDA率%": to_pct(r.get("ebitdaMargin")),
            "ROE%": to_pct(r.get("returnOnEquity")),
            "ROA%": to_pct(r.get("returnOnAssets")),
            "ROIC%": to_pct(r.get("returnOnInvestedCapital")),
            "流動比": r.get("currentRatio"), "速動比": r.get("quickRatio"),
            "D/E": r.get("debtToEquityRatio") or r.get("debtEquityRatio"),
            "P/E": r.get("priceToEarningsRatio"),
            "P/B": r.get("priceToBookRatio"),
            "P/S": r.get("priceToSalesRatio"),
            "P/FCF": r.get("priceToFreeCashFlowRatio"),
        } for r in ratios]
        sheets["比率10Y"] = pd.DataFrame(rows).sort_values("年", ascending=False)

    # === 分析師預估 ===
    est = data.get("estimates") or []
    if est:
        rows = [{
            "年": (e.get("date") or "")[:4],
            "營收Avg": e.get("revenueAvg"), "營收High": e.get("revenueHigh"),
            "營收Low": e.get("revenueLow"),
            "EBITDA": e.get("ebitdaAvg"), "淨利": e.get("netIncomeAvg"),
            "EPS Avg": e.get("epsAvg"), "EPS High": e.get("epsHigh"), "EPS Low": e.get("epsLow"),
            "分析師#(rev)": e.get("numAnalystsRevenue"),
            "分析師#(eps)": e.get("numAnalystsEps"),
        } for e in est]
        sheets["分析師預估"] = pd.DataFrame(rows).sort_values("年")

    # === 產品/地理結構 ===
    for key, tag, colname in [("prod_seg","產品結構","產品"), ("geo_seg","地理結構","地區")]:
        seg = data.get(key) or []
        if seg:
            rows = []
            for r in seg:
                year = r.get("fiscalYear")
                d = r.get("data",{}) if isinstance(r.get("data"), dict) else {}
                total = sum(v for v in d.values() if isinstance(v,(int,float)) and v)
                for item, rev in sorted(d.items(), key=lambda x: x[1] or 0, reverse=True):
                    if rev:
                        rows.append({"年":year, colname:item, "營收":rev,
                                     "占比%": round(rev/total*100, 1) if total>0 else None})
            if rows: sheets[tag] = pd.DataFrame(rows)

    # === 內部人統計 (近 8 季) ===
    ist = data.get("insider_stats") or []
    if ist:
        rows = [{
            "年季": f"{r.get('year')}Q{r.get('quarter')}",
            "買筆": r.get("acquiredTransactions"), "賣筆": r.get("disposedTransactions"),
            "買賣比": r.get("acquiredDisposedRatio"),
            "買量": r.get("totalAcquired"), "賣量": r.get("totalDisposed"),
        } for r in ist[:8]]
        sheets["內部人統計"] = pd.DataFrame(rows)

    # === 內部人明細 ===
    ins = data.get("insider_search") or []
    if ins:
        rows = [{
            "日期": r.get("transactionDate"), "姓名": r.get("reportingName"),
            "身分": r.get("typeOfOwner"), "類型": r.get("transactionType"),
            "股數": r.get("securitiesTransacted"), "價格": r.get("price"),
            "持有": r.get("securitiesOwned"),
        } for r in ins]
        sheets["內部人明細"] = pd.DataFrame(rows)

    # === 國會 ===
    cong = []
    for tx in (data.get("senate") or [])[:15]:
        cong.append({"院":"參議院", "議員":f"{tx.get('firstName','')} {tx.get('lastName','')}".strip(),
                     "州":tx.get("district"), "日期":tx.get("transactionDate"),
                     "類型":tx.get("type"), "金額":tx.get("amount"),
                     "所有人":tx.get("owner")})
    for tx in (data.get("house") or [])[:15]:
        cong.append({"院":"眾議院", "議員":f"{tx.get('firstName','')} {tx.get('lastName','')}".strip(),
                     "州":tx.get("district"), "日期":tx.get("transactionDate"),
                     "類型":tx.get("type"), "金額":tx.get("amount"),
                     "所有人":tx.get("owner")})
    if cong:
        sheets["國會交易"] = pd.DataFrame(cong).sort_values("日期", ascending=False)

    # === 5% 大戶 ===
    bo = data.get("beneficial") or []
    if bo:
        rows = [{
            "日期":r.get("filingDate"), "報告人":r.get("nameOfReportingPerson"),
            "類型":r.get("typeOfReportingPerson"),
            "持股":r.get("amountBeneficiallyOwned"), "占比%":r.get("percentOfClass"),
        } for r in bo[:15]]
        sheets["5%大戶"] = pd.DataFrame(rows)

    # === 評等歷史 ===
    gh = data.get("grades_hist") or []
    if gh:
        rows = [{"日期":r.get("date"), "SB":r.get("analystRatingsStrongBuy"),
                 "B":r.get("analystRatingsBuy"), "H":r.get("analystRatingsHold"),
                 "S":r.get("analystRatingsSell"), "SS":r.get("analystRatingsStrongSell")}
                for r in gh]
        sheets["評等歷史"] = pd.DataFrame(rows)

    # === 財報 vs 預期 ===
    earn = data.get("earnings") or []
    if earn:
        rows = [{"日期":e.get("date"), "EPS實":e.get("epsActual"), "EPS預":e.get("epsEstimated"),
                 "營收實":e.get("revenueActual"), "營收預":e.get("revenueEstimated")}
                for e in earn]
        sheets["財報vs預期"] = pd.DataFrame(rows)

    # === Owner Earnings ===
    oe = data.get("owner_e") or []
    if oe:
        rows = [{"日期":r.get("date"),
                 "年季":f"{r.get('fiscalYear','')}{r.get('period','')}",
                 "OE":r.get("ownersEarnings"), "OE/股":r.get("ownersEarningsPerShare"),
                 "維護 CapEx":r.get("maintenanceCapex"), "成長 CapEx":r.get("growthCapex")}
                for r in oe]
        sheets["OwnerEarnings"] = pd.DataFrame(rows)

    # === 員工數 ===
    emp = data.get("emp_hist") or []
    if emp:
        rows = [{"日期":r.get("periodOfReport"), "員工數":r.get("employeeCount")}
                for r in emp]
        sheets["員工數"] = pd.DataFrame(rows)

    # === 高管 ===
    kx = data.get("key_executives") or []
    if kx:
        rows = [{"職稱":r.get("title"), "姓名":r.get("name"),
                 "薪酬":r.get("pay"), "生年":r.get("yearBorn")} for r in kx]
        sheets["高管"] = pd.DataFrame(rows)

    # === 同業 ===
    peers = data.get("peers") or []
    if peers:
        rows = [{"代號":p.get("symbol"), "公司":p.get("companyName"),
                 "股價":p.get("price"), "市值":p.get("mktCap")} for p in peers]
        sheets["同業"] = pd.DataFrame(rows)

    # === 股利 ===
    div = data.get("dividends") or []
    if div:
        rows = [{"除息":r.get("date"), "發放":r.get("paymentDate"),
                 "股利":r.get("dividend"), "殖利率%":r.get("yield")} for r in div]
        sheets["股利歷史"] = pd.DataFrame(rows)

    # === 技術面 ===
    def to_df(x, val_col, name):
        if not x: return pd.DataFrame()
        return pd.DataFrame([{"date":r.get("date"), name:r.get(val_col)} for r in x])
    a = to_df(data.get("sma50"), "sma", "MA50")
    b = to_df(data.get("sma200"), "sma", "MA200")
    c = to_df(data.get("rsi"), "rsi", "RSI")
    if not a.empty:
        merged = a
        for d in [b, c]:
            if not d.empty: merged = merged.merge(d, on="date", how="outer")
        sheets["技術面"] = merged.sort_values("date")

    # === 新聞 ===
    news = data.get("news") or []
    if news:
        rows = [{"日期":r.get("publishedDate"), "來源":r.get("publisher"),
                 "標題":r.get("title"), "URL":r.get("url")} for r in news]
        sheets["新聞"] = pd.DataFrame(rows)

    # === SEC 8-K ===
    sec = data.get("sec_8k") or []
    if sec:
        rows = [{"日期":r.get("filingDate"), "類型":r.get("formType"),
                 "連結":r.get("finalLink")} for r in sec]
        sheets["SEC_8K"] = pd.DataFrame(rows)

    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with pd.ExcelWriter(dst, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return sheets


def _calc_entry_us(price, ma200, yr_low, tp_low):
    """3:1 計算, 返回 dict 或 None"""
    if not (price and ma200 and yr_low and tp_low and tp_low > 0):
        return None
    sl = max(ma200, yr_low * 1.03)
    if ma200 >= price:
        sl = yr_low * 1.05
    if tp_low <= sl:
        return None
    max_entry = (tp_low + 3 * sl) / 4
    if price <= sl: verdict = "🚨 已破止損(SL 失守)"
    elif price <= max_entry: verdict = "🟢 進場區(≥3:1)"
    elif price <= max_entry * 1.05: verdict = "🟡 接近門檻(<5%)"
    elif price <= max_entry * 1.15: verdict = "🟠 稍高(5-15%)"
    else: verdict = "🔴 追高風險(>15%)"
    ratio = round((tp_low - price) / (price - sl), 2) if price > sl else None
    dist_pct = round((price / max_entry - 1) * 100, 1)
    return {"sl": round(sl, 2), "tp": round(tp_low, 2), "max_entry": round(max_entry, 2),
            "verdict": verdict, "ratio": ratio, "dist_pct": dist_pct}


def _us_signals(data):
    """內部人 4Q 買賣比 + 國會 90d 淨"""
    from datetime import datetime, timedelta
    out = {"insider_ratio": None, "insider_signal": None,
           "cong_buy": 0, "cong_sell": 0, "cong_signal": None}
    ist = data.get("insider_stats") or []
    if ist:
        recent = ist[:4]
        buy = sum(float(r.get("totalAcquired") or 0) for r in recent)
        sell = sum(float(r.get("totalDisposed") or 0) for r in recent)
        if sell > 0:
            r = round(buy / sell, 2)
            out["insider_ratio"] = r
            if r < 0.05:   out["insider_signal"] = "🔴 極端賣"
            elif r < 0.3:  out["insider_signal"] = "🟠 賣壓"
            elif r > 5:    out["insider_signal"] = "🟢 極端買"
            elif r > 2:    out["insider_signal"] = "🟡 買方"
            else:          out["insider_signal"] = "⚪ 中性"
    cutoff = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
    for tx in (data.get("senate") or []) + (data.get("house") or []):
        d = tx.get("transactionDate", "") or ""
        if d < cutoff: continue
        t = (tx.get("type") or "").lower()
        if "purchase" in t or t == "buy": out["cong_buy"] += 1
        elif "sale" in t or t == "sell": out["cong_sell"] += 1
    net = out["cong_buy"] - out["cong_sell"]
    total = out["cong_buy"] + out["cong_sell"]
    if total == 0: out["cong_signal"] = None
    elif net >= 3:  out["cong_signal"] = "🟢 強買"
    elif net >= 1:  out["cong_signal"] = "🟡 買方"
    elif net <= -3: out["cong_signal"] = "🔴 強賣"
    elif net <= -1: out["cong_signal"] = "🟠 賣方"
    else:           out["cong_signal"] = "⚪ 中性"
    return out


def _event_radar_us_md(entry, sig):
    lines = ["\n## 🚨 事件雷達 (3:1 + 內部人 + 國會)\n",
             "| 訊號 | 值 | 判讀 |", "|---|---:|---|"]
    has = False
    if entry:
        lines.append(f"| **3:1 判讀** | 距 Max Entry {entry['dist_pct']:+.1f}% | **{entry['verdict']}** |")
        has = True
    if sig["insider_ratio"] is not None:
        lines.append(f"| 內部人 4Q 買賣比 | {sig['insider_ratio']} | {sig['insider_signal']} |")
        has = True
    if sig["cong_signal"]:
        lines.append(f"| 國會 90d (買/賣筆數) | {sig['cong_buy']}/{sig['cong_sell']} | {sig['cong_signal']} |")
        has = True
    return "\n".join(lines) + "\n" if has else ""


def _summary_action_us(entry, sig):
    if not entry:
        return ""
    md = f"""
---

## 🎯 綜合判讀 / 事件行動

### 進場策略 (3:1 派)
- **限價買**: ${entry['max_entry']}
- **止損**: ${entry['sl']} (跌破止損, 重新評估)
- **目標**: ${entry['tp']}
- **當前判讀**: {entry['verdict']}
"""
    warnings = []
    if sig["insider_signal"] and "🔴" in sig["insider_signal"]:
        warnings.append(f"內部人 {sig['insider_signal']} (4Q 買賣比 {sig['insider_ratio']})")
    if sig["cong_signal"] and "🔴" in sig["cong_signal"]:
        warnings.append(f"國會 {sig['cong_signal']} (買 {sig['cong_buy']} / 賣 {sig['cong_sell']})")
    if warnings:
        md += "\n### 🚨 短期警訊\n" + "\n".join(f"- {w}" for w in warnings) + "\n"
    md += """
### 撤退訊號 (任一觸發即檢討)
1. EPS 連 2 季 QoQ 下滑
2. 內部人 4Q 買賣比跌破 0.1 (內部人賣壓極端)
3. 分析師目標中位下調 > 10%
4. 股價跌破 SL

⚠️ 3:1 SL 用 MA200/52w低 較高者, TP 用分析師最保守 targetLow, 實戰請自行核對技術支撐
"""
    return md


def _per_band_us(price, data):
    """PER × TTM EPS 換算價. 返回 (md, df) 或 (None, None)"""
    rttm = first(data.get("ratios_ttm")) or {}
    kmttm = first(data.get("km_ttm")) or {}
    eps = None
    for k in ("netIncomePerShareTTM", "earningsPerShareTTM", "epsTTM"):
        v = kmttm.get(k) or rttm.get(k)
        if v: eps = float(v); break
    if not eps and rttm.get("priceToEarningsRatioTTM") and price:
        pe = float(rttm["priceToEarningsRatioTTM"])
        if pe > 0: eps = price / pe
    if not eps:
        inc = data.get("income") or []
        if inc and inc[0].get("eps"): eps = float(inc[0]["eps"])
    if not eps or eps <= 0 or not price:
        return None, None
    bands = [10, 15, 20, 25, 30, 40]
    rows = []
    for pe in bands:
        tgt = round(eps * pe, 2)
        diff = round((price / tgt - 1) * 100, 1)
        if diff < -20:   v = "🟢 大幅低估"
        elif diff < -5:  v = "🟢 低估"
        elif diff < 5:   v = "🟡 合理"
        elif diff < 20:  v = "🟠 高估"
        else:            v = "🔴 大幅高估"
        rows.append({"PER 倍": f"{pe}x", "對應價": tgt, "距現價 %": diff, "判讀": v})
    df = pd.DataFrame(rows)
    cur_pe = round(price / eps, 1)
    md = f"""
## 📊 PER 換算價 (基於 TTM EPS)

- TTM EPS: **${eps:.2f}**
- 現價 ${price} → 對應 **PER {cur_pe}x**

| PER | 對應價 | 距現價 % | 判讀 |
|---:|---:|---:|---|
"""
    for r in rows:
        s = '+' if r['距現價 %'] > 0 else ''
        md += f"| {r['PER 倍']} | ${r['對應價']} | {s}{r['距現價 %']}% | {r['判讀']} |\n"
    md += """
⚠️ 假設 EPS 維持. 若 EPS 加速成長 → 高倍數也合理; 若衰退 → 低倍數才安全
💡 3:1 (技術) vs PER (估值) 兩個視角常會給出不同答案 — 用來 cross check
"""
    return md, df


def _entry_section_us(price, ma200, yr_low, tp_low, n):
    """3:1 盈虧比入場計算段落 (April1Stock 公式)"""
    e = _calc_entry_us(price, ma200, yr_low, tp_low)
    if not e:
        if price and ma200 and yr_low and tp_low and tp_low > 0:
            sl = max(ma200, yr_low * 1.03)
            return f"\n## 🎯 3:1 入場計算\n\n⚠️ TP (${n(tp_low)}) ≤ SL (${n(sl)}) — 分析師目標已跌破支撐, 公式無法計算\n"
        return ""
    ratio_s = n(e["ratio"]) if e["ratio"] is not None else "—"
    return f"""
## 🎯 3:1 入場計算 (April1Stock 公式)

Max Entry = (TP + 3×SL) / 4 → 現價 ≤ Max Entry 才符合 3:1 盈虧比

| 項目 | 值 | 說明 |
|---|---:|---|
| **判讀** | **{e['verdict']}** | |
| SL 止損 | ${n(e['sl'])} | MA200 或 52w低×1.03 取高 |
| TP 目標 | ${n(e['tp'])} | 分析師最保守 targetLow |
| **Max Entry** | **${n(e['max_entry'])}** | 現價 ≤ 此值才買 |
| 現價 | ${n(price)} | 距 Max Entry {'+' if e['dist_pct'] > 0 else ''}{e['dist_pct']}% |
| 實際盈虧比 | {ratio_s} | 目標 ≥ 3.0 |
"""


def build_md(sym, data, dst):
    """Markdown 精華"""
    prof  = first(data.get("profile")) or {}
    quote = first(data.get("quote")) or {}
    pc    = first(data.get("price_change")) or {}
    scores= first(data.get("scores")) or {}
    tgt   = first(data.get("target_consensus")) or {}
    grades= first(data.get("grades_consensus")) or {}
    dcf   = first(data.get("dcf")) or {}
    dcf_l = first(data.get("dcf_lev")) or {}
    kmttm = first(data.get("km_ttm")) or {}
    rttm  = first(data.get("ratios_ttm")) or {}
    fg    = first(data.get("fin_growth")) or {}
    price = quote.get("price")

    def n(v, d=2):
        try: return f"{float(v):,.{d}f}"
        except: return "—"
    def p(v):
        try: return f"{float(v)*100:.1f}%" if v is not None else "—"
        except: return "—"
    def sig(v):
        if v is None: return ""
        return "+" if v > 0 else ""

    mkt_cap = quote.get("marketCap")
    mkt_cap_s = f"${mkt_cap/1e9:,.1f}B" if mkt_cap else "—"
    dcf_diff = round((dcf.get("dcf")/price-1)*100,1) if price and dcf.get("dcf") else None
    tgt_diff = round((tgt.get("targetMedian")/price-1)*100,1) if price and tgt.get("targetMedian") else None

    md = f"""# 🔍 {sym} — {prof.get('companyName','')}

**{prof.get('sector','—')} / {prof.get('industry','—')}** | CEO: {prof.get('ceo','—')} | Employees: {prof.get('fullTimeEmployees','—')} | IPO: {prof.get('ipoDate','—')}

> {(prof.get('description') or '')[:300]}...

---

## 💰 報價 & 估值

| 指標 | 值 | 指標 | 值 |
|---|---:|---|---:|
| 股價 | **${n(price)}** | 市值 | {mkt_cap_s} |
| 52w 高/低 | {n(quote.get('yearHigh'))} / {n(quote.get('yearLow'))} | Beta | {n(prof.get('beta'))} |
| MA50 / MA200 | {n(quote.get('priceAvg50'))} / {n(quote.get('priceAvg200'))} | — | — |
| **P/E TTM** | **{n(rttm.get('priceToEarningsRatioTTM'))}** | Fwd P/E | {n(rttm.get('forwardPriceToEarningsGrowthRatioTTM'))} |
| PEG | {n(rttm.get('priceToEarningsGrowthRatioTTM'))} | P/S | {n(rttm.get('priceToSalesRatioTTM'))} |
| **EV/EBITDA** | **{n(kmttm.get('evToEBITDATTM'))}** | EV/FCF | {n(kmttm.get('evToFreeCashFlowTTM'))} |
| **DCF** | **${n(dcf.get('dcf'))}** | Levered DCF | ${n(dcf_l.get('dcf'))} |
| **DCF vs 現價** | **{sig(dcf_diff)}{dcf_diff}%** | Graham # | {n(kmttm.get('grahamNumberTTM'))} |
| FCF Yield | {p(kmttm.get('freeCashFlowYieldTTM'))} | Earnings Yield | {p(kmttm.get('earningsYieldTTM'))} |

## 📈 漲跌

| 期 | % | 期 | % |
|---|---:|---|---:|
| 1D | {n(pc.get('1D'))}% | 1Y | {n(pc.get('1Y'))}% |
| 1M | {n(pc.get('1M'))}% | 3Y | {n(pc.get('3Y'))}% |
| 3M | {n(pc.get('3M'))}% | 5Y | {n(pc.get('5Y'))}% |
| YTD | {n(pc.get('ytd'))}% | 10Y | {n(pc.get('10Y'))}% |

## 🎯 分析師

- **目標價中位: ${n(tgt.get('targetMedian'))}** ({sig(tgt_diff)}{tgt_diff}% vs 現價)
- 高 / 低: ${n(tgt.get('targetHigh'))} / ${n(tgt.get('targetLow'))}
- 華爾街評等 **{grades.get('consensus','—')}**: SB {grades.get('strongBuy',0)} | B {grades.get('buy',0)} | H {grades.get('hold',0)} | S {grades.get('sell',0)}
{_event_radar_us_md(_calc_entry_us(price, quote.get('priceAvg200'), quote.get('yearLow'), tgt.get('targetLow')), _us_signals(data))}
{_entry_section_us(price, quote.get('priceAvg200'), quote.get('yearLow'), tgt.get('targetLow'), n)}
{_per_band_us(price, data)[0] or ''}
## 🏥 體質

| 項目 | 值 | 判讀 |
|---|---:|---|
| Altman Z | **{n(scores.get('altmanZScore'))}** | {'🟢 安全' if (scores.get('altmanZScore') or 0) >= 3 else '🟠 中性' if (scores.get('altmanZScore') or 0) >= 1.8 else '🔴 風險'} |
| Piotroski | **{scores.get('piotroskiScore','—')}/9** | {'🟢 強' if (scores.get('piotroskiScore') or 0) >= 7 else '🟡 中性' if (scores.get('piotroskiScore') or 0) >= 4 else '🔴 弱'} |
| ROE | **{p(rttm.get('returnOnEquityTTM'))}** | ROIC {p(kmttm.get('returnOnInvestedCapitalTTM'))} |
| 毛利率 | **{p(rttm.get('grossProfitMarginTTM'))}** | 淨利率 {p(rttm.get('netProfitMarginTTM'))} |

## 📊 成長率

| 期間 | 營收/股 | 淨利/股 | OCF/股 | 股利/股 |
|---|---:|---:|---:|---:|"""

    def cs(c, N):
        v = cumul_to_cagr(c, N)
        return f"{v}%" if v is not None else "—"
    for label, prefix, N in [("10Y","tenY",10), ("5Y","fiveY",5), ("3Y","threeY",3)]:
        md += f"\n| {label} CAGR | {cs(fg.get(f'{prefix}RevenueGrowthPerShare'), N)} | {cs(fg.get(f'{prefix}NetIncomeGrowthPerShare'), N)} | {cs(fg.get(f'{prefix}OperatingCFGrowthPerShare'), N)} | {cs(fg.get(f'{prefix}DividendperShareGrowthPerShare'), N)} |"
    md += f"\n| 1Y YoY | {p(fg.get('revenueGrowth'))} | {p(fg.get('netIncomeGrowth'))} | {p(fg.get('operatingCashFlowGrowth'))} | {p(fg.get('dividendsPerShareGrowth'))} |"

    # 產品結構
    ps = data.get("prod_seg") or []
    if ps:
        latest = ps[0]
        d = latest.get("data",{}) if isinstance(latest.get("data"), dict) else {}
        total = sum(v for v in d.values() if isinstance(v,(int,float)) and v)
        md += f"\n\n## 🛍️ 產品結構 (FY{latest.get('fiscalYear','')})\n\n| 產品 | 營收 | 占比 |\n|---|---:|---:|"
        for prod, rev in sorted(d.items(), key=lambda x: x[1] or 0, reverse=True)[:8]:
            if rev: md += f"\n| {prod} | ${rev/1e9:,.1f}B | {rev/total*100:.1f}% |"

    # 地理結構
    gs = data.get("geo_seg") or []
    if gs:
        latest = gs[0]
        d = latest.get("data",{}) if isinstance(latest.get("data"), dict) else {}
        total = sum(v for v in d.values() if isinstance(v,(int,float)) and v)
        md += f"\n\n## 🌍 地理 (FY{latest.get('fiscalYear','')})\n\n| 地區 | 營收 | 占比 |\n|---|---:|---:|"
        for r, rev in sorted(d.items(), key=lambda x: x[1] or 0, reverse=True)[:8]:
            if rev: md += f"\n| {r} | ${rev/1e9:,.1f}B | {rev/total*100:.1f}% |"

    # 內部人
    ist = data.get("insider_stats") or []
    if ist:
        latest = ist[0]
        md += f"\n\n## 👤 內部人 (最新季 {latest.get('year')}Q{latest.get('quarter')})\n- 買 {latest.get('acquiredTransactions',0)} 筆 / 賣 {latest.get('disposedTransactions',0)} 筆\n- 買賣比: **{latest.get('acquiredDisposedRatio','—')}**\n"

    # 國會
    sen = data.get("senate") or []
    hou = data.get("house") or []
    if sen or hou:
        md += "\n## 🏛️ 國會交易 (最近)\n\n"
        for tx in (sen[:3] + hou[:3]):
            side = "🟢買" if tx.get("type","").lower() in ("purchase","buy") else "🔴賣"
            md += f"- {tx.get('transactionDate','')} {side} {tx.get('firstName','')} {tx.get('lastName','')} ({tx.get('amount','')})\n"

    # 同業
    peers = data.get("peers") or []
    if peers:
        md += "\n## 👥 同業\n\n"
        for p in peers[:8]:
            md += f"- **{p.get('symbol')}** {p.get('companyName')} — ${p.get('price','—')} / ${(p.get('mktCap') or 0)/1e9:.1f}B\n"

    # 新聞
    news = data.get("news") or []
    if news:
        md += "\n## 📰 最新新聞\n\n"
        for x in news[:5]:
            md += f"- [{x.get('title','')}]({x.get('url','')}) — *{x.get('publisher','')} ({(x.get('publishedDate') or '')[:10]})*\n"

    md += _summary_action_us(
        _calc_entry_us(price, quote.get('priceAvg200'), quote.get('yearLow'), tgt.get('targetLow')),
        _us_signals(data))

    md += f"\n\n---\n\n_資料源: FMP • 生成 {datetime.now().strftime('%Y-%m-%d %H:%M')}_\n"

    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(md)


def process(sym):
    print(f"\n=== {sym} 深度研究 ===")
    t0 = time.time()
    data = fetch_all(sym)
    ok = sum(1 for v in data.values() if v)
    print(f"  抓 {ok}/{len(data)} 端點 ({time.time()-t0:.1f}s)")
    xlsx_dst = os.path.join(OUT_DIR, f"深度研究_{sym}.xlsx")
    md_dst = os.path.join(OUT_DIR, f"深度研究_{sym}.md")
    sheets = build_xlsx(sym, data, xlsx_dst)
    build_md(sym, data, md_dst)
    print(f"  → {xlsx_dst}  ({len(sheets)} 分頁)")
    print(f"  → {md_dst}")


def main():
    if not KEY: print("⚠️ 需 FMP_API_KEY"); sys.exit(1)
    # 支援兩種傳法: sys.argv 或 env TICKER
    args = [a.upper() for a in sys.argv[1:] if a.strip()]
    if not args:
        env = os.environ.get("TICKER", "").upper().strip()
        if env: args = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
    if not args:
        print("⚠️ 未指定代號: python us_stock_deepdive.py NVDA"); sys.exit(1)
    for sym in args: process(sym)


if __name__ == "__main__":
    main()
