# -*- coding: utf-8 -*-
"""
單檔美股深度分析 us_stock_deepdive.py
=======================================================================
用法:
  python us_stock_deepdive.py NVDA
  python us_stock_deepdive.py NVDA META GOOG    # 多檔

輸出:
  data/deepdive/{SYM}_深度.xlsx  含以下分頁:
    [10年營收EPS] - 10年年度數字+YoY+營收拐點旗標
    [近4季財報]   - 最新季營收/EPS/毛利/淨利 vs 去年同期
    [長期報酬]    - 10/5/3/1y 含息年化報酬率
    [近期走勢]    - 6/3/1m 周線(每週收盤)+ 漲跌%
    [買進訊號]    - 當前 vs MA50/MA200/RSI14/距52週高低
  console 一頁總結
"""
import os
import sys
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
OUT_DIR = "data/deepdive"
os.makedirs(OUT_DIR, exist_ok=True)


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1)); continue
        if r.status_code != 200:
            return None
        try: return r.json()
        except: return None
    return None


def fmt_pct(v):
    if v is None or pd.isna(v): return "-"
    return f"{v:+.1f}%"


def annual_revenue_eps(sym):
    """10年年度營收 / 毛利 / 淨利 / EPS"""
    inc = get("income-statement", symbol=sym, period="annual", limit=10) or []
    if not inc: return None
    rows = []
    for r in reversed(inc):
        rows.append({
            "年度": r.get("calendarYear") or r.get("date", "")[:4],
            "營收(億)": round((r.get("revenue") or 0) / 1e8, 1),
            "毛利率%": round((r.get("grossProfitRatio") or 0) * 100, 1),
            "淨利率%": round((r.get("netIncomeRatio") or 0) * 100, 1),
            "EPS稀釋": round(r.get("epsdiluted") or 0, 2),
        })
    df = pd.DataFrame(rows)
    # YoY
    df["營收YoY%"] = df["營收(億)"].pct_change().mul(100).round(1)
    df["EPS YoY%"] = df["EPS稀釋"].pct_change().mul(100).round(1)
    # 拐點:近 3 年連續正成長後出現負,或連續負後出現正
    yoy = df["營收YoY%"].fillna(0).values
    flag = [""] * len(yoy)
    for i in range(3, len(yoy)):
        if yoy[i] < -5 and all(v > 5 for v in yoy[i-2:i]):
            flag[i] = "⚠️ 由盛轉衰"
        if yoy[i] > 5 and all(v < -5 for v in yoy[i-2:i]):
            flag[i] = "🟢 反轉向上"
    df["營收拐點"] = flag
    return df


def latest_quarters(sym):
    """近 4 季 vs 去年同期"""
    inc = get("income-statement", symbol=sym, period="quarter", limit=8) or []
    if not inc: return None
    rows = []
    for r in inc:
        rows.append({
            "季": r.get("period",""), "日期": r.get("date",""),
            "營收(億)": round((r.get("revenue") or 0) / 1e8, 1),
            "毛利率%": round((r.get("grossProfitRatio") or 0) * 100, 1),
            "淨利率%": round((r.get("netIncomeRatio") or 0) * 100, 1),
            "EPS稀釋": round(r.get("epsdiluted") or 0, 2),
        })
    df = pd.DataFrame(rows).sort_values("日期").reset_index(drop=True)
    # 季 YoY (跟 4 季前同期比)
    if len(df) >= 5:
        df["營收 YoY%"] = ((df["營收(億)"] - df["營收(億)"].shift(4)) / df["營收(億)"].shift(4) * 100).round(1)
        df["EPS YoY%"] = ((df["EPS稀釋"] - df["EPS稀釋"].shift(4)) / df["EPS稀釋"].shift(4) * 100).round(1)
    return df.tail(4).reset_index(drop=True)


def fetch_prices(sym, years=11):
    """抓 ~11 年 EOD,使用 adjClose(已含股息再投資)"""
    today = datetime.now(timezone.utc).date()
    start = today - timedelta(days=int(years*365.25))
    data = get("historical-price-eod/full", symbol=sym,
               **{"from": start.isoformat(), "to": today.isoformat()})
    if not data: return None
    # FMP 回 dict 含 'historical' 或直接 list
    hist = data.get("historical") if isinstance(data, dict) else data
    if not hist: return None
    df = pd.DataFrame(hist)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
    # adjClose 若無 → 用 close
    if "adjClose" not in df.columns:
        df["adjClose"] = df["close"]
    return df[["date","close","adjClose"]]


def long_term_returns(prices):
    """10/5/3/1y 含息年化報酬率"""
    if prices is None or len(prices) == 0: return None
    last = prices.iloc[-1]
    last_date = last["date"]
    last_p = last["adjClose"]
    rows = []
    for label, yrs in [("10y", 10), ("5y", 5), ("3y", 3), ("1y", 1)]:
        target = last_date - pd.Timedelta(days=int(yrs*365.25))
        sub = prices[prices["date"] <= target]
        if len(sub) == 0: continue
        p0 = sub.iloc[-1]["adjClose"]
        total = (last_p / p0 - 1) * 100
        cagr = ((last_p / p0) ** (1/yrs) - 1) * 100
        rows.append({"區間": label, "起始日": sub.iloc[-1]["date"].date(),
                     "起始價(調)": round(p0,2), "現價(調)": round(last_p,2),
                     "總報酬%": round(total,1), "年化%": round(cagr,2)})
    return pd.DataFrame(rows)


def recent_weekly(prices):
    """6/3/1m 周線(每週五收盤)"""
    if prices is None: return None
    last_date = prices["date"].max()
    rows = []
    for label, days in [("6m", 180), ("3m", 90), ("1m", 30)]:
        sub = prices[prices["date"] >= (last_date - pd.Timedelta(days=days))].copy()
        if len(sub) == 0: continue
        sub = sub.set_index("date").resample("W-FRI").last().dropna(subset=["close"]).reset_index()
        if len(sub) < 2: continue
        chg = (sub.iloc[-1]["close"] / sub.iloc[0]["close"] - 1) * 100
        rows.append({"區間": label, "週數": len(sub),
                     "起始": round(sub.iloc[0]["close"], 2),
                     "現價": round(sub.iloc[-1]["close"], 2),
                     "區間漲跌%": round(chg, 1)})
    return pd.DataFrame(rows)


def buy_signal(prices):
    """買進訊號:現價 vs MA50/MA200/RSI14/52w高低"""
    if prices is None or len(prices) < 200: return None
    p = prices["close"].values
    cur = p[-1]
    ma50  = p[-50:].mean()
    ma200 = p[-200:].mean()

    # RSI14
    delta = pd.Series(p).diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = (-delta.clip(upper=0)).rolling(14).mean()
    rs = gain / loss
    rsi = (100 - 100 / (1 + rs)).iloc[-1]

    # 52 週高低
    p52 = prices[prices["date"] >= prices["date"].max() - pd.Timedelta(days=365)]["close"]
    hi52 = p52.max(); lo52 = p52.min()

    # 拐點判讀
    signals = []
    if cur < ma50 < ma200:
        signals.append("🔴 空頭排列")
    elif cur > ma50 > ma200:
        signals.append("🟢 多頭排列")
    if rsi < 30:
        signals.append(f"🟢 超賣(RSI {rsi:.0f})")
    elif rsi > 70:
        signals.append(f"🔴 超買(RSI {rsi:.0f})")
    if cur < lo52 * 1.05:
        signals.append("🟢 接近52週低")
    if cur > hi52 * 0.97:
        signals.append("🔴 接近52週高")

    return pd.DataFrame([{
        "現價": round(cur,2), "MA50": round(ma50,2), "MA200": round(ma200,2),
        "距MA50%": round((cur/ma50-1)*100,1),
        "距MA200%": round((cur/ma200-1)*100,1),
        "RSI14": round(rsi,1),
        "52週高": round(hi52,2), "距52高%": round((cur/hi52-1)*100,1),
        "52週低": round(lo52,2), "距52低%": round((cur/lo52-1)*100,1),
        "綜合訊號": " | ".join(signals) if signals else "中性"
    }])


def analyze(sym):
    print(f"\n{'='*60}\n  {sym} 深度分析\n{'='*60}")
    rev = annual_revenue_eps(sym)
    qtr = latest_quarters(sym)
    prices = fetch_prices(sym)
    ret = long_term_returns(prices)
    wk = recent_weekly(prices)
    sig = buy_signal(prices)

    # console
    if rev is not None:
        print(f"\n[10 年年度營收 EPS]"); print(rev.to_string(index=False))
    if qtr is not None:
        print(f"\n[近 4 季]"); print(qtr.to_string(index=False))
    if ret is not None:
        print(f"\n[10/5/3/1y 含息年化報酬]"); print(ret.to_string(index=False))
    if wk is not None:
        print(f"\n[近期周線走勢]"); print(wk.to_string(index=False))
    if sig is not None:
        print(f"\n[買進訊號]"); print(sig.to_string(index=False))

    # 寫 Excel
    out = f"{OUT_DIR}/{sym}_深度.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        if rev is not None: rev.to_excel(xw, sheet_name="10年營收EPS", index=False)
        if qtr is not None: qtr.to_excel(xw, sheet_name="近4季財報", index=False)
        if ret is not None: ret.to_excel(xw, sheet_name="長期報酬", index=False)
        if wk is not None:  wk.to_excel(xw, sheet_name="近期走勢", index=False)
        if sig is not None: sig.to_excel(xw, sheet_name="買進訊號", index=False)
    print(f"\n→ 已輸出 {out}")


def main():
    if not KEY: print("⚠️ 未設 FMP_API_KEY"); return
    syms = [s.upper() for s in sys.argv[1:]]
    if not syms: print("用法: python us_stock_deepdive.py SYM1 SYM2 ..."); return
    for s in syms: analyze(s)


if __name__ == "__main__":
    main()
