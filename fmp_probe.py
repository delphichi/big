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

# ────────────────────────────────────────────────────────────────
# 端點分類:每個端點記 (endpoint, params_example, 說明)
# ────────────────────────────────────────────────────────────────
ENDPOINTS = [
    # === 報價 / 概況 ===
    ("📈 報價即時", [
        ("quote",                {"symbol": TICKER}, "即時報價 + PE/EPS/52w高低"),
        ("quote-short",          {"symbol": TICKER}, "簡版報價"),
        ("price-target-summary", {"symbol": TICKER}, "分析師目標價彙總"),
        ("price-target-consensus", {"symbol": TICKER}, "目標價共識"),
        ("price-target-latest-news", {"symbol": TICKER, "limit": 3}, "最新目標價變動"),
        ("grades-consensus",     {"symbol": TICKER}, "分析師評級共識"),
        ("upgrades-downgrades",  {"symbol": TICKER, "limit": 5}, "升降評紀錄"),
    ]),
    # === 公司基本 ===
    ("🏢 公司資訊", [
        ("profile",              {"symbol": TICKER}, "公司簡介 + 產業 + CEO + 員工數"),
        ("market-capitalization", {"symbol": TICKER}, "市值"),
        ("company-notes",        {"symbol": TICKER}, "公司債券發行紀錄"),
        ("employee-count",       {"symbol": TICKER, "limit": 3}, "員工數歷年"),
        ("executive-compensation", {"symbol": TICKER, "limit": 3}, "高管薪酬"),
    ]),
    # === 三大報表 ===
    ("📊 三大報表(年)", [
        ("income-statement",     {"symbol": TICKER, "period": "annual", "limit": 3}, "損益表 - 營收/成本/淨利/EPS"),
        ("balance-sheet-statement", {"symbol": TICKER, "period": "annual", "limit": 3}, "資產負債表"),
        ("cash-flow-statement",  {"symbol": TICKER, "period": "annual", "limit": 3}, "現金流量表"),
    ]),
    ("📊 三大報表(季)", [
        ("income-statement",     {"symbol": TICKER, "period": "quarter", "limit": 4}, "季損益"),
        ("balance-sheet-statement", {"symbol": TICKER, "period": "quarter", "limit": 4}, "季資產負債"),
        ("cash-flow-statement",  {"symbol": TICKER, "period": "quarter", "limit": 4}, "季現金流"),
    ]),
    # === 成長率(FMP 已經幫你算)===
    ("📈 成長率現成", [
        ("income-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "損益表 YoY 成長"),
        ("balance-sheet-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "BS 成長"),
        ("cash-flow-statement-growth", {"symbol": TICKER, "period": "annual", "limit": 3}, "CF 成長"),
        ("financial-growth",     {"symbol": TICKER, "period": "annual", "limit": 3}, "綜合財務成長(3y/5y/10y CAGR 都有)"),
    ]),
    # === 估值 / 關鍵比率 ===
    ("💎 估值 / 比率", [
        ("ratios",               {"symbol": TICKER, "period": "annual", "limit": 3}, "估值比率(PE/PB/ROE...)"),
        ("ratios-ttm",           {"symbol": TICKER}, "TTM 比率"),
        ("key-metrics",          {"symbol": TICKER, "period": "annual", "limit": 3}, "關鍵指標(每股FCF/PEG/債比...)"),
        ("key-metrics-ttm",      {"symbol": TICKER}, "TTM 關鍵指標"),
        ("enterprise-values",    {"symbol": TICKER, "period": "annual", "limit": 3}, "企業價值 EV"),
        ("owner-earnings",       {"symbol": TICKER, "limit": 3}, "Buffett 式擁有者盈餘"),
    ]),
    # === 預估 / 分析師 ===
    ("🔮 分析師預估", [
        ("analyst-estimates",    {"symbol": TICKER, "period": "annual", "limit": 3}, "分析師 EPS / 營收預估"),
        ("ratings-snapshot",     {"symbol": TICKER}, "評級快照"),
        ("ratings-historical",   {"symbol": TICKER, "limit": 5}, "歷史評級"),
    ]),
    # === 財報事件 ===
    ("📅 財報日曆", [
        ("earnings",             {"symbol": TICKER, "limit": 4}, "歷史財報日期 + 數字"),
        ("earnings-surprises",   {"symbol": TICKER, "limit": 4}, "EPS 預期 vs 實際"),
        ("earnings-transcript",  {"symbol": TICKER, "limit": 1}, "法說會逐字稿"),
    ]),
    # === 股利 / 拆股 ===
    ("💰 股利 / 拆股", [
        ("dividends",            {"symbol": TICKER, "limit": 5}, "歷史股利"),
        ("splits",               {"symbol": TICKER, "limit": 5}, "歷史拆股"),
    ]),
    # === 持股 / 內部 ===
    ("🏦 持股 / 內部人", [
        ("institutional-ownership-list", {"symbol": TICKER, "limit": 5}, "機構持股清單"),
        ("insider-trading",      {"symbol": TICKER, "limit": 5}, "內部人交易"),
        ("insider-roster",       {"symbol": TICKER}, "高管/董事名單"),
        ("etf-holdings",         {"symbol": TICKER, "limit": 5}, "ETF 持有此檔(反查)"),
    ]),
    # === 歷史價量 ===
    ("📉 歷史價量", [
        ("historical-price-eod/light", {"symbol": TICKER, "from": "2025-01-01"}, "日 K 線(輕量)"),
        ("historical-price-eod/full",  {"symbol": TICKER, "from": "2025-12-01"}, "日 K + adj close + vol"),
        ("historical-price-eod/dividend-adjusted", {"symbol": TICKER, "from": "2025-12-01"}, "股利調整後"),
        ("historical-market-capitalization", {"symbol": TICKER, "from": "2025-01-01", "limit": 5}, "歷史市值"),
    ]),
    # === 評等 / 排名 ===
    ("⭐ 評等指標", [
        ("financial-scores",     {"symbol": TICKER}, "FMP 自家財務評分(0-5 Piotroski 等)"),
        ("rating",               {"symbol": TICKER}, "FMP 綜合評等"),
        ("rating-historical",    {"symbol": TICKER, "limit": 5}, "歷史評等"),
        ("discounted-cash-flow", {"symbol": TICKER}, "DCF 估值"),
        ("levered-dcf",          {"symbol": TICKER}, "Levered DCF"),
        ("custom-dcf",           {"symbol": TICKER}, "客製 DCF"),
    ]),
    # === SEC 文件 ===
    ("📁 SEC 文件", [
        ("sec-filings-search/symbol", {"symbol": TICKER, "limit": 5}, "近期 SEC filings"),
        ("sec-filings-financials", {"symbol": TICKER, "limit": 3}, "財報文件"),
    ]),
    # === 整體市場 ===
    ("🌍 整體市場(全市場, 無需 symbol)", [
        ("biggest-gainers",      {}, "今日漲幅前列"),
        ("biggest-losers",       {}, "今日跌幅前列"),
        ("most-actives",         {}, "今日成交量前列"),
        ("market-hours",         {}, "市場開盤狀態"),
        ("sector-performance-snapshot", {"date": "2025-12-01"}, "各產業表現"),
        ("industry-performance-snapshot", {"date": "2025-12-01"}, "細產業表現"),
        ("treasury-rates",       {"from": "2025-12-01"}, "美國公債利率"),
        ("economic-indicators",  {"name": "CPI"}, "經濟指標(CPI/GDP/失業率)"),
    ]),
    # === Crypto / Forex / Commodities ===
    ("🌐 加密 / 外匯 / 商品", [
        ("cryptocurrency-list",  {}, "加密幣清單"),
        ("forex-list",           {}, "外匯對清單"),
        ("commodities-list",     {}, "商品清單"),
    ]),
]


def fetch(endpoint, params):
    url = f"{BASE}/{endpoint}"
    params = dict(params, apikey=KEY)
    try:
        r = requests.get(url, params=params, timeout=TIMEOUT)
        if r.status_code == 200:
            try:
                j = r.json()
                return r.status_code, j
            except Exception as e:
                return r.status_code, f"json parse err: {e}"
        return r.status_code, (r.text[:200] if r.text else "")
    except Exception as e:
        return -1, f"req err: {e}"


def summarize(data):
    """簡化 JSON 顯示, 只看 keys + 第一筆"""
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
            return f"list[{len(data)}] × dict({len(keys)} fields):\n   欄位: {keys[:20]}{'...' if len(keys)>20 else ''}\n   第一筆例: {json.dumps(first, ensure_ascii=False, default=str)[:300]}"
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
            if status == 200:
                ok += 1
                out(f"   ⇢ {summarize(data)}")
            else:
                bad += 1
                out(f"   ⇢ {summarize(data)}")
            time.sleep(0.1)  # 禮貌間隔

    out(f"\n\n=== 總計: ✅ {ok}  ❌ {bad}  / {ok+bad} ===")

    if OUT:
        os.makedirs(os.path.dirname(OUT) or ".", exist_ok=True)
        with open(OUT, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print(f"\n→ 寫檔: {OUT}")


if __name__ == "__main__":
    main()
