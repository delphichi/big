# -*- coding: utf-8 -*-
"""
美股六維交叉 us_six_dim.py
=======================================================================
合併 3 個 dashboard 資料 + 新算報酬, 產出 95 檔六維評分:

  1. 成長     ← data/美股95檔_加速度分類.xlsx
  2. 估值     ← data/美股_全景儀表板.xlsx (DCF + LDCF)
  3. 體質     ← data/美股_全景儀表板.xlsx (Altman + Piotroski)
  4. 內部人   ← data/美股_全景儀表板.xlsx (4Q 買賣比)
  5. 國會     ← data/美股_全景儀表板.xlsx (90d + 強訊號)
  6. 報酬     ← FMP /stock-price-change (1y/3y/5y) + vs SPY 超額

跑法: FMP_API_KEY=xxx python us_six_dim.py
輸出: data/美股95檔_六維交叉.xlsx
"""
import os, time, requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
WATCHLIST_FILE = "data/watchlist_us.txt"
WORKERS = int(os.environ.get("WORKERS", "6"))

GROWTH_SRC = "data/美股95檔_加速度分類.xlsx"
DASH_SRC = "data/美股_全景儀表板.xlsx"
DST = "data/美股95檔_六維交叉.xlsx"


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip().upper() for t in env.replace(",", " ").split() if t.strip()]
        return list(dict.fromkeys([t for t in toks if t and not t.startswith("#")]))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip().upper() for t in line.split() if t.strip())
        return list(dict.fromkeys(toks))
    return "NVDA AVGO MSFT".split()


def get(endpoint, **params):
    if not KEY: return None
    params["apikey"] = KEY
    for _ in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=15)
            if r.status_code == 429: time.sleep(2); continue
            if r.status_code != 200: return None
            return r.json()
        except Exception: time.sleep(1)
    return None


def fetch_change(sym):
    """抓 stock-price-change (1D/5D/1M/3M/6M/YTD/1Y/3Y/5Y/10Y)"""
    d = get("stock-price-change", symbol=sym)
    if not d or not isinstance(d, list): return sym, {}
    r = d[0]
    return sym, {
        "1m": r.get("1M"), "3m": r.get("3M"), "6m": r.get("6M"),
        "1y": r.get("1Y"), "3y": r.get("3Y"), "5y": r.get("5Y"), "10y": r.get("10Y"),
    }


def six_score(r):
    """六維評分"""
    s = 0; flags = []

    # 1. 成長 (加速度分類)
    cat = str(r.get("分類", ""))
    if "加速器" in cat: s += 2; flags.append("🚀加速")
    elif "高飛" in cat: s += 1; flags.append("✈️高飛")
    elif "失速" in cat: s -= 2; flags.append("💀失速")
    elif "減速" in cat: s -= 1; flags.append("📉減速")

    # 2. 估值 (DCF 差%, 已用 v2 過濾規則)
    d = r.get("DCF差%")
    if pd.notna(d):
        if 30 < d <= 100: s += 2; flags.append("💎深度低估")
        elif 10 < d <= 30: s += 1
        elif -30 <= d < -10: s -= 1
        elif -100 < d < -30: s -= 2; flags.append("🔴深度高估")

    # 3. 體質 (Altman 對非金融 + Piotroski)
    fin = str(r.get("產業","")).strip() in {"Financial Services","Financials","Real Estate"}
    z = r.get("AltmanZ")
    if not fin and pd.notna(z):
        if z >= 3: s += 1
        elif z < 1.8: s -= 2; flags.append("💀破產風險")
    p = r.get("Piotroski")
    if pd.notna(p):
        if p >= 8: s += 1; flags.append("🟢體質強")
        elif p <= 3: s -= 1

    # 4. 內部人 (4Q 合計)
    ir = r.get("內部人買賣比")
    if pd.notna(ir):
        if ir >= 1: s += 1; flags.append("👤內部人淨買")
        elif ir < 0.1: s -= 1; flags.append("⚠️內部人大賣")

    # 5. 國會強訊號
    sig = r.get("國會強訊號")
    if sig == "強買": s += 1; flags.append("🏛️🟢國會強買")
    elif sig == "強賣": s -= 1; flags.append("🏛️🔴國會強賣")

    # 6. 報酬 (vs SPY 1y 超額)
    e = r.get("1y超額%")
    if pd.notna(e):
        if e > 50: s += 2; flags.append("🏆暴贏 SPY")
        elif e > 0: s += 1
        elif e < -30: s -= 2; flags.append("📉跌輸 SPY")

    # 評等 A + 品質 90+ 加分
    if str(r.get("評等","")) == "A" and (r.get("品質總分") or 0) >= 90:
        s += 1; flags.append("🏅A品90+")

    return s, " ".join(flags) if flags else "—"


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY, 無法抓報酬"); return
    codes = load_watchlist()
    print(f"美股六維交叉 — {len(codes)} 檔")

    # ─── 載入 3 個現成資料 ───
    grow = pd.read_excel(GROWTH_SRC)
    print(f"成長分類: {len(grow)} 檔")

    dash = pd.read_excel(DASH_SRC, sheet_name="總覽")
    print(f"全景儀表板: {len(dash)} 檔")

    # 統一代號 str
    grow["代號"] = grow["代號"].astype(str)
    dash["代號"] = dash["代號"].astype(str)

    # ─── 抓報酬 (95 calls) ───
    print(f"抓 stock-price-change (含 SPY 大盤)...")
    syms_with_spy = codes + ["SPY"]
    returns = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_change, s): s for s in syms_with_spy}
        done = 0
        for fut in as_completed(futs):
            sym, d = fut.result()
            if d: returns[sym] = d
            done += 1
            if done % 20 == 0: print(f"  [{done}/{len(syms_with_spy)}]")

    spy = returns.get("SPY", {})
    print(f"SPY 報酬: 1y={spy.get('1y')} 3y={spy.get('3y')} 5y={spy.get('5y')} 10y={spy.get('10y')}")

    # ─── 組裝主表 ───
    master = grow[["代號","名稱","評等","品質","營收10y","營收5y","營收3y","淨利3y","分類"]].rename(
        columns={"品質":"品質總分"}).copy()

    # merge full_dashboard 欄位
    dash_keep = [c for c in ["代號","產業","當前股價","DCF差%","LDCF差%","AltmanZ","Piotroski",
                              "OE/股","OE殖利率%","主產品1","中國營收%","台灣營收%",
                              "內部人買賣比","國會90d買","國會90d賣","國會強訊號","綜合分","訊號"] if c in dash.columns]
    master = master.merge(dash[dash_keep].rename(columns={"訊號":"全景訊號","綜合分":"全景分"}),
                          on="代號", how="left")

    # 加報酬 + 超額
    for sym in master["代號"]:
        pass
    def gp(s, k):
        d = returns.get(s, {})
        return d.get(k)
    master["1y報酬%"] = master["代號"].apply(lambda s: gp(s, "1y"))
    master["3y報酬%"] = master["代號"].apply(lambda s: gp(s, "3y"))
    master["5y報酬%"] = master["代號"].apply(lambda s: gp(s, "5y"))
    master["10y報酬%"] = master["代號"].apply(lambda s: gp(s, "10y"))
    # 對 SPY 超額
    spy_1y = spy.get("1y"); spy_3y = spy.get("3y"); spy_5y = spy.get("5y")
    if spy_1y is not None: master["1y超額%"] = master["1y報酬%"] - spy_1y
    if spy_3y is not None: master["3y超額%"] = master["3y報酬%"] - spy_3y
    if spy_5y is not None: master["5y超額%"] = master["5y報酬%"] - spy_5y

    # ─── 六維評分 ───
    master["六維分"], master["六維訊號"] = zip(*master.apply(six_score, axis=1))
    master = master.sort_values("六維分", ascending=False)

    # 欄位順序
    front = [c for c in ["代號","名稱","產業","評等","品質總分","分類","當前股價",
                          "DCF差%","AltmanZ","Piotroski","內部人買賣比","國會強訊號",
                          "1y報酬%","1y超額%","3y報酬%","3y超額%","5y超額%",
                          "六維分","六維訊號","全景訊號"] if c in master.columns]
    rest = [c for c in master.columns if c not in front]
    master = master[front + rest]

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        master.to_excel(xw, sheet_name="總覽", index=False)
        master.head(20).to_excel(xw, sheet_name="TOP20", index=False)
        master.tail(15).to_excel(xw, sheet_name="BOTTOM15", index=False)
        # SPY 對照
        pd.DataFrame([spy]).to_excel(xw, sheet_name="SPY大盤", index=False)

    print(f"\n→ {DST}")
    print(f"\n=== TOP 15 (六維分) ===")
    show = ["代號","名稱","評等","分類","DCF差%","Piotroski","1y超額%","3y超額%","六維分","六維訊號"]
    print(master[[c for c in show if c in master.columns]].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
