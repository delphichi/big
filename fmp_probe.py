# -*- coding: utf-8 -*-
"""
FMP 端點偵錯 fmp_probe.py
=======================================================================
列出 FMP /stable/ 常用端點 + 每個拿一筆數據看欄位

跑法:
  FMP_API_KEY=xxx python fmp_probe.py            # 預設用 NVDA
  PROBE_TICKER=AAPL python fmp_probe.py
  PROBE_TICKER=NVDA OUT=data/fmp_probe.txt python fmp_probe.py

輸出: 終端 + (可選) 寫檔
"""
import os, json, sys, time
import requests

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
TICKER = os.environ.get("PROBE_TICKER", "NVDA").upper()
OUT = os.environ.get("OUT", "")
TIMEOUT = 15

ENDPOINTS = [
    # === 報價 / 概況 ===
    ("📈 報價即時", [
        ("quote",                {"symbol": TICKER}, "即時報價 + PE/EPS/52w 高低"),
        ("quote-short",          {"symbol": TICKER}, "簡版報價"),
        ("aftermarket-quote",    {"symbol": TICKER}, "盤後報價"),
        ("stock-price-change",   {"symbol": TICKER}, "1d/5d/1m/3m/1y 漲跌"),
    ]),
    # === 公司基本 ===
    ("🏢 公司資訊", [
        ("profile",              {"symbol": TICKER}, "公司簡介 + 產業 + CEO + 員工數"),
        ("market-capitalization", {"symbol": TICKER}, "市值"),
        ("historical-market-capitalization", {"symbol": TICKER, "from": "2025-01-01", "limit": 5}, "歷史市值"),
        ("shares-float",         {"symbol": TICKER}, "流通股數 / 內部人持有比例"),
        ("key-executives",       {"symbol": TICKER}, "高管名單"),
        ("employee-count",       {"symbol": TICKER, "limit": 3}, "員工數歷年"),
        ("historical-employee-count", {"symbol": TICKER, "limit": 5}, "員工數歷年(更詳細)"),
        ("executive-compensation-benchmark", {"symbol": TICKER}, "高管薪酬"),
        ("stock-peers",          {"symbol": TICKER}, "FMP 推算的同業公司"),
    ]),
    # === 三大報表 ===
    ("📊 三大報表(年)", [
        ("income-statement",     {"symbol": TICKER, "period": "annual", "limit": 3}, "損益表"),
        ("balance-sheet-statement", {"symbol": TICKER, "period": "annual", "limit": 3}, "資產負債表"),
        ("cash-flow-statement",  {"symbol": TICKER, "period": "annual", "limit": 3}, "現金流量表"),
        ("income-statement-ttm", {"symbol": TICKER}, "TTM 損益"),
        ("cash-flow-statement-ttm", {"symbol": TICKER}, "TTM 現金流"),
        ("balance-sheet-statement-ttm", {"symbol": TICKER}, "TTM 資產負債"),
        ("latest-financial-statements", {"symbol": TICKER}, "最新三表"),
        ("financial-reports-dates", {"symbol": TICKER}, "10-K/10-Q 日期"),
    ]),
    # === 成長率 ===
    ("📈 成長率現成", [
        ("income-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "損益表 YoY"),
        ("balance-sheet-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "BS YoY"),
        ("cash-flow-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "CF YoY"),
        ("financial-growth",     {"symbol": TICKER, "period": "annual", "limit": 3}, "綜合 3y/5y/10y CAGR ⭐"),
    ]),
    # === 營收結構(地理 / 產品)===
    ("🌍 營收細分", [
        ("revenue-geographic-segmentation", {"symbol": TICKER, "period": "annual", "limit": 3}, "地理區域 ⭐"),
        ("revenue-product-segmentation", {"symbol": TICKER, "period": "annual", "limit": 3}, "產品線 ⭐"),
    ]),
    # === 估值 / 關鍵比率 ===
    ("💎 估值 / 比率", [
        ("ratios",               {"symbol": TICKER, "period": "annual", "limit": 3}, "估值比率 PE/PB/ROE"),
        ("ratios-ttm",           {"symbol": TICKER}, "TTM 比率"),
        ("key-metrics",          {"symbol": TICKER, "period": "annual", "limit": 3}, "關鍵指標"),
        ("key-metrics-ttm",      {"symbol": TICKER}, "TTM 關鍵指標"),
        ("enterprise-values",    {"symbol": TICKER, "period": "annual", "limit": 3}, "企業價值 EV"),
        ("owner-earnings",       {"symbol": TICKER, "limit": 3}, "Buffett 式 Owner Earnings ⭐"),
        ("financial-scores",     {"symbol": TICKER}, "FMP 自家評分(Piotroski/Altman)"),
        ("discounted-cash-flow", {"symbol": TICKER}, "DCF 估值"),
        ("levered-discounted-cash-flow", {"symbol": TICKER}, "Levered DCF"),
    ]),
    # === 預估 / 分析師 ===
    ("🔮 分析師預估", [
        ("analyst-estimates",    {"symbol": TICKER, "period": "annual", "limit": 3}, "EPS / 營收預估"),
        ("price-target-summary", {"symbol": TICKER}, "目標價 summary"),
        ("price-target-consensus", {"symbol": TICKER}, "目標價共識"),
        ("ratings-snapshot",     {"symbol": TICKER}, "FMP 自家評等"),
        ("ratings-historical",   {"symbol": TICKER, "limit": 5}, "歷史評等"),
        ("grades",               {"symbol": TICKER, "limit": 5}, "華爾街評等"),
        ("grades-consensus",     {"symbol": TICKER}, "華爾街評等共識"),
        ("grades-historical",    {"symbol": TICKER, "limit": 5}, "華爾街評等歷史"),
        ("tipranks-symbol-summary", {"symbol": TICKER}, "TipRanks 評等(可能付費)"),
    ]),
    # === 財報事件 ===
    ("📅 財報日曆", [
        ("earnings",             {"symbol": TICKER, "limit": 4}, "歷史財報日期"),
        ("earnings-calendar",    {}, "全市場財報日曆"),
        ("earning-call-transcript", {"symbol": TICKER, "year": 2025}, "法說會逐字稿 ⭐"),
        ("earning-call-transcript-dates", {"symbol": TICKER}, "法說會日期"),
    ]),
    # === 股利 / 拆股 ===
    ("💰 股利 / 拆股", [
        ("dividends",            {"symbol": TICKER, "limit": 5}, "歷史股利"),
        ("dividends-calendar",   {}, "全市場除息日曆"),
        ("splits",               {"symbol": TICKER, "limit": 5}, "歷史拆股"),
        ("splits-calendar",      {}, "全市場拆股日曆"),
    ]),
    # === 持股 / 內部人 / 國會 ===
    ("🏦 持股 / 內部人 / 國會", [
        ("institutional-ownership/latest", {"symbol": TICKER, "limit": 5}, "機構持股最新"),
        ("institutional-ownership/extract", {"symbol": TICKER, "year": 2025, "quarter": 3}, "機構持股展開"),
        ("institutional-ownership/symbol-positions-summary", {"symbol": TICKER, "year": 2025, "quarter": 3}, "機構持股 summary"),
        ("insider-trading/search", {"symbol": TICKER, "limit": 5}, "內部人交易 ⭐"),
        ("insider-trading/statistics", {"symbol": TICKER}, "內部人交易統計"),
        ("acquisition-of-beneficial-ownership", {"symbol": TICKER}, "5%+ 持股變動"),
        ("senate-trades",        {"symbol": TICKER}, "參議員交易此檔 ⭐"),
        ("house-trades",         {"symbol": TICKER}, "眾議員交易此檔 ⭐"),
        ("etf/holdings",         {"symbol": TICKER}, "(輸 ETF 代號才有效)"),
        ("etf/asset-exposure",   {"symbol": TICKER}, "哪些 ETF 持有這檔 ⭐"),
    ]),
    # === 歷史價量 ===
    ("📉 歷史價量", [
        ("historical-price-eod/light", {"symbol": TICKER, "from": "2025-12-01"}, "日 K(輕)"),
        ("historical-price-eod/full",  {"symbol": TICKER, "from": "2025-12-01"}, "日 K(完整 + adj close)"),
        ("historical-price-eod/dividend-adjusted", {"symbol": TICKER, "from": "2025-12-01"}, "股利調整"),
        ("historical-chart/1day",      {"symbol": TICKER, "from": "2025-12-01"}, "日線"),
        ("historical-chart/1hour",     {"symbol": TICKER, "from": "2025-12-29"}, "小時線"),
    ]),
    # === 技術指標 ===
    ("📊 技術指標", [
        ("technical-indicators/sma", {"symbol": TICKER, "periodLength": 50, "timeframe": "1day", "from": "2025-12-01"}, "MA50"),
        ("technical-indicators/ema", {"symbol": TICKER, "periodLength": 20, "timeframe": "1day", "from": "2025-12-01"}, "EMA20"),
        ("technical-indicators/rsi", {"symbol": TICKER, "periodLength": 14, "timeframe": "1day", "from": "2025-12-01"}, "RSI14"),
        ("technical-indicators/adx", {"symbol": TICKER, "periodLength": 14, "timeframe": "1day", "from": "2025-12-01"}, "ADX"),
        ("technical-indicators/standarddeviation", {"symbol": TICKER, "periodLength": 20, "timeframe": "1day", "from": "2025-12-01"}, "標準差"),
    ]),
    # === SEC ===
    ("📁 SEC 文件", [
        ("sec-filings-search/symbol", {"symbol": TICKER, "limit": 5}, "近期 SEC filings"),
        ("sec-filings-financials", {"symbol": TICKER, "limit": 3}, "財報文件"),
        ("sec-filings-8k",       {"symbol": TICKER, "limit": 3}, "8-K 重大事件"),
        ("sec-profile",          {"symbol": TICKER}, "SEC profile"),
        ("financial-reports-json", {"symbol": TICKER, "year": 2024, "period": "FY"}, "完整財報 JSON"),
    ]),
    # === ESG ===
    ("🌿 ESG", [
        ("esg-disclosures",      {"symbol": TICKER}, "ESG 揭露"),
        ("esg-ratings",          {"symbol": TICKER}, "ESG 評等"),
        ("esg-benchmark",        {"year": 2024}, "ESG 產業基準"),
    ]),
    # === 整體市場(無 symbol)===
    ("🌍 整體市場", [
        ("biggest-gainers",      {}, "今日漲幅前列"),
        ("biggest-losers",       {}, "今日跌幅前列"),
        ("most-actives",         {}, "今日成交量前列"),
        ("market-hours",         {"exchange": "NASDAQ"}, "市場開盤狀態"),
        ("sector-performance-snapshot", {"date": "2025-12-01"}, "各產業表現 ⭐"),
        ("industry-performance-snapshot", {"date": "2025-12-01"}, "細產業表現 ⭐"),
        ("sector-pe-snapshot",   {"date": "2025-12-01"}, "各產業 PE ⭐"),
        ("industry-pe-snapshot", {"date": "2025-12-01"}, "細產業 PE ⭐"),
        ("historical-sector-pe", {"sector": "Technology", "from": "2025-01-01"}, "歷史產業 PE ⭐"),
        ("historical-industry-pe", {"industry": "Semiconductors", "from": "2025-01-01"}, "歷史細產業 PE"),
        ("historical-sector-performance", {"sector": "Technology", "from": "2025-01-01"}, "歷史產業表現"),
        ("historical-industry-performance", {"industry": "Semiconductors", "from": "2025-01-01"}, "歷史細產業表現"),
        ("treasury-rates",       {"from": "2025-12-01"}, "美國公債利率"),
        ("economic-indicators",  {"name": "CPI"}, "經濟指標(CPI/GDP/失業率)⭐"),
        ("economic-calendar",    {"from": "2025-12-15"}, "經濟事件日曆"),
        ("market-risk-premium",  {}, "市場風險溢酬"),
    ]),
    # === 指數成分 ===
    ("📑 指數成分", [
        ("sp500-constituent",    {}, "S&P 500 成分"),
        ("nasdaq-constituent",   {}, "NASDAQ 成分"),
        ("dowjones-constituent", {}, "Dow 30 成分"),
        ("historical-sp500-constituent", {}, "歷史 S&P 500 異動"),
    ]),
    # === IPO / M&A / 募資 ===
    ("🚀 IPO / M&A", [
        ("ipos-calendar",        {}, "IPO 日曆"),
        ("ipos-disclosure",      {}, "IPO 揭露"),
        ("ipos-prospectus",      {}, "IPO 招股書"),
        ("mergers-acquisitions-latest", {}, "最新 M&A"),
        ("mergers-acquisitions-search", {"name": "NVIDIA"}, "搜 M&A"),
        ("fundraising-latest",   {}, "最新私募"),
        ("crowdfunding-offerings-latest", {}, "群眾募資"),
    ]),
    # === News ===
    ("📰 新聞", [
        ("news/stock",           {"symbols": TICKER, "limit": 3}, "個股新聞"),
        ("news/general-latest",  {"limit": 3}, "綜合新聞"),
        ("news/press-releases",  {"symbols": TICKER, "limit": 3}, "公司發稿"),
        ("fmp-articles",         {"limit": 3}, "FMP 自家分析文"),
    ]),
    # === COT 期貨持倉 ===
    ("📑 COT 期貨持倉", [
        ("commitment-of-traders-list", {}, "COT 商品清單"),
        ("commitment-of-traders-report", {"symbol": "GC", "from": "2025-12-01"}, "黃金 COT"),
        ("commitment-of-traders-analysis", {"symbol": "GC", "from": "2025-12-01"}, "黃金 COT 分析 ⭐"),
    ]),
    # === Crypto / Forex / Commodities ===
    ("🌐 其他資產", [
        ("cryptocurrency-list",  {}, "加密幣清單"),
        ("forex-list",           {}, "外匯清單"),
        ("commodities-list",     {}, "商品清單"),
        ("batch-commodity-quotes", {}, "全商品即時"),
        ("batch-crypto-quotes",  {}, "全加密即時"),
        ("batch-forex-quotes",   {}, "全外匯即時"),
    ]),
    # === Screener ===
    ("🔍 篩選器", [
        ("company-screener",     {"marketCapMoreThan": 100000000000, "limit": 5}, "公司篩選(市值 > 1000 億)"),
        ("stock-list",           {}, "全市場代號"),
        ("actively-trading-list", {}, "活躍交易股"),
        ("delisted-companies",   {}, "下市公司"),
    ]),
]


def fetch(endpoint, params):
    url = f"{BASE}/{endpoint}"
    params = dict(params, apikey=KEY)
    try:
        r = requests.get(url, params=params, timeout=TIMEOUT)
        if r.status_code == 200:
            try: return r.status_code, r.json()
            except Exception as e: return r.status_code, f"json parse err: {e}"
        return r.status_code, (r.text[:200] if r.text else "")
    except Exception as e:
        return -1, f"req err: {e}"


def summarize(data):
    if data is None: return "None"
    if isinstance(data, str): return f"[非 JSON] {data[:120]}"
    if isinstance(data, dict):
        if "Error Message" in data or "error" in data:
            return f"❌ ERR: {data}"
        keys = list(data.keys())[:10]
        return f"dict({len(data)} keys): {keys}"
    if isinstance(data, list):
        if not data: return "[] (空陣列)"
        first = data[0]
        if isinstance(first, dict):
            keys = list(first.keys())
            return f"list[{len(data)}] × dict({len(keys)} fields):\n   欄位: {keys[:25]}{'...' if len(keys)>25 else ''}\n   第一筆例: {json.dumps(first, ensure_ascii=False, default=str)[:350]}"
        return f"list[{len(data)}] 第一個: {first}"
    return str(data)[:200]


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY, 無法呼叫"); sys.exit(1)

    lines = []
    def out(s):
        print(s); lines.append(s)

    out(f"=== FMP 端點偵錯 ({TICKER}) ===\n")
    ok = bad = 0
    for group, eps in ENDPOINTS:
        out(f"\n{'='*70}\n{group}\n{'='*70}")
        for endpoint, params, desc in eps:
            status, data = fetch(endpoint, params)
            mark = "✅" if status == 200 else "❌"
            param_s = " ".join(f"{k}={v}" for k,v in params.items() if k != "apikey")
            out(f"\n{mark} [{status}] /{endpoint}  {param_s}")
            out(f"   📝 {desc}")
            out(f"   ⇢ {summarize(data)}")
            if status == 200: ok += 1
            else: bad += 1
            time.sleep(0.1)

    out(f"\n\n=== 總計: ✅ {ok}  ❌ {bad}  / {ok+bad} ===")

    if OUT:
        os.makedirs(os.path.dirname(OUT) or ".", exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"\n→ 寫檔: {OUT}")


if __name__ == "__main__":
    main()
