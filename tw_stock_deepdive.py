# -*- coding: utf-8 -*-
"""
單檔台股深度分析 tw_stock_deepdive.py
=======================================================================
用法:
  python tw_stock_deepdive.py 2330
  python tw_stock_deepdive.py 2330 3017 6139

輸出:
  data/tw_deepdive/{CODE}_深度.xlsx 含 5 分頁:
    [10年營收EPS] - 年度營收/EPS + YoY + 拐點旗標
    [近12月營收]  - 月營收 + YoY + MoM
    [近4季財報]   - 最新季 vs 去年同期
    [長期報酬]    - 10/5/3/1y 含息年化(用 TaiwanStockPriceAdj)
    [近期走勢]    - 6/3/1m 周線
    [買進訊號]    - MA50/MA200/RSI14/52週高低
"""
import os
import sys
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
OUT_DIR = "data/tw_deepdive"
os.makedirs(OUT_DIR, exist_ok=True)


def get(dataset, **params):
    params["dataset"] = dataset
    if TOKEN: params["token"] = TOKEN
    for attempt in range(3):
        try:
            r = requests.get(BASE, params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 402: time.sleep(5); continue
        if r.status_code != 200: return None
        j = r.json()
        if j.get("status") != 200: return None
        return j.get("data", [])
    return None


def annual_revenue_eps(code):
    """10 年年度營收+EPS:從季財報加總"""
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=int(11*365.25))).isoformat()
    fin = get("TaiwanStockFinancialStatements", data_id=code, start_date=start)
    if not fin: return None
    df = pd.DataFrame(fin)
    if df.empty: return None
    # 抓營收 & EPS
    rev = df[df["type"] == "Revenue"][["date","value"]].rename(columns={"value":"營收"})
    eps = df[df["type"] == "EPS"][["date","value"]].rename(columns={"value":"EPS"})
    gp  = df[df["type"] == "GrossProfit"][["date","value"]].rename(columns={"value":"毛利"})
    ni  = df[df["type"] == "IncomeAfterTaxes"][["date","value"]].rename(columns={"value":"稅後淨利"})
    m = rev
    for d in (eps, gp, ni):
        if not d.empty: m = m.merge(d, on="date", how="left")
    m["date"] = pd.to_datetime(m["date"])
    m["年度"] = m["date"].dt.year
    m["季"] = m["date"].dt.quarter
    # 加總成年度
    yr = m.groupby("年度").agg({"營收":"sum","EPS":"sum","毛利":"sum","稅後淨利":"sum"})
    yr["營收(億)"] = (yr["營收"]/1e8).round(1)
    yr["毛利率%"] = (yr["毛利"]/yr["營收"]*100).round(1)
    yr["淨利率%"] = (yr["稅後淨利"]/yr["營收"]*100).round(1)
    yr["EPS"] = yr["EPS"].round(2)
    yr["營收YoY%"] = yr["營收(億)"].pct_change().mul(100).round(1)
    yr["EPS YoY%"] = yr["EPS"].pct_change().mul(100).round(1)
    yoy = yr["營收YoY%"].fillna(0).values
    flag = [""] * len(yoy)
    for i in range(3, len(yoy)):
        if yoy[i] < -5 and all(v > 5 for v in yoy[i-2:i]): flag[i] = "⚠️ 由盛轉衰"
        if yoy[i] > 5 and all(v < -5 for v in yoy[i-2:i]): flag[i] = "🟢 反轉向上"
    yr["營收拐點"] = flag
    yr = yr.reset_index()[["年度","營收(億)","毛利率%","淨利率%","EPS","營收YoY%","EPS YoY%","營收拐點"]]
    return yr


def monthly_revenue(code):
    """近 12 月月營收 + YoY/MoM"""
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=540)).isoformat()
    r = get("TaiwanStockMonthRevenue", data_id=code, start_date=start)
    if not r: return None
    df = pd.DataFrame(r).sort_values("date").reset_index(drop=True)
    df["營收(億)"] = (df["revenue"]/1e8).round(1)
    df["YoY%"] = pd.to_numeric(df.get("revenue_year"), errors="coerce").round(1)
    df["MoM%"] = pd.to_numeric(df.get("revenue_month"), errors="coerce").round(1)
    df["年月"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m")
    return df[["年月","營收(億)","YoY%","MoM%"]].tail(12).reset_index(drop=True)


def latest_quarters(code):
    """近 4 季 vs 去年同期"""
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=730)).isoformat()
    fin = get("TaiwanStockFinancialStatements", data_id=code, start_date=start)
    if not fin: return None
    df = pd.DataFrame(fin)
    rev = df[df["type"]=="Revenue"][["date","value"]].rename(columns={"value":"營收"})
    eps = df[df["type"]=="EPS"][["date","value"]].rename(columns={"value":"EPS"})
    gp  = df[df["type"]=="GrossProfit"][["date","value"]].rename(columns={"value":"毛利"})
    ni  = df[df["type"]=="IncomeAfterTaxes"][["date","value"]].rename(columns={"value":"稅後淨利"})
    m = rev
    for d in (eps, gp, ni):
        if not d.empty: m = m.merge(d, on="date", how="left")
    m["date"] = pd.to_datetime(m["date"])
    m = m.sort_values("date").reset_index(drop=True)
    m["營收(億)"] = (m["營收"]/1e8).round(1)
    m["毛利率%"] = (m["毛利"]/m["營收"]*100).round(1)
    m["淨利率%"] = (m["稅後淨利"]/m["營收"]*100).round(1)
    m["EPS"] = m["EPS"].round(2)
    m["營收YoY%"] = ((m["營收(億)"]-m["營收(億)"].shift(4))/m["營收(億)"].shift(4)*100).round(1)
    m["EPS YoY%"] = ((m["EPS"]-m["EPS"].shift(4))/m["EPS"].shift(4)*100).round(1)
    m["季"] = m["date"].dt.strftime("%Y-Q%q").str.replace("Q1","Q1").str.replace("Q2","Q2").str.replace("Q3","Q3").str.replace("Q4","Q4")
    m["季"] = m["date"].dt.year.astype(str) + "Q" + m["date"].dt.quarter.astype(str)
    return m[["季","營收(億)","毛利率%","淨利率%","EPS","營收YoY%","EPS YoY%"]].tail(4).reset_index(drop=True)


def fetch_prices(code, years=11):
    """抓 ~11 年 EOD,用 TaiwanStockPriceAdj(已含股息再投資)"""
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=int(years*365.25))).isoformat()
    r = get("TaiwanStockPriceAdj", data_id=code, start_date=start)
    if not r:
        r = get("TaiwanStockPrice", data_id=code, start_date=start)
    if not r: return None
    df = pd.DataFrame(r).sort_values("date").reset_index(drop=True)
    df["date"] = pd.to_datetime(df["date"])
    if "close" not in df.columns: return None
    df["adjClose"] = df["close"]  # PriceAdj 已經是調整後
    return df[["date","close","adjClose"]]


def long_term_returns(prices):
    if prices is None or len(prices)==0: return None
    last = prices.iloc[-1]
    rows = []
    for label, yrs in [("10y",10),("5y",5),("3y",3),("1y",1)]:
        tgt = last["date"] - pd.Timedelta(days=int(yrs*365.25))
        sub = prices[prices["date"] <= tgt]
        if len(sub)==0 or sub.iloc[-1]["adjClose"] <= 0: continue
        p0 = sub.iloc[-1]["adjClose"]
        total = (last["adjClose"]/p0 - 1) * 100
        cagr = ((last["adjClose"]/p0) ** (1/yrs) - 1) * 100
        rows.append({"區間":label, "起始日":sub.iloc[-1]["date"].date(),
                     "起始價(調)":round(p0,2), "現價(調)":round(last["adjClose"],2),
                     "總報酬%":round(total,1), "年化%":round(cagr,2)})
    return pd.DataFrame(rows)


def recent_weekly(prices):
    if prices is None: return None
    last = prices["date"].max()
    rows = []
    for label, days in [("6m",180),("3m",90),("1m",30)]:
        sub = prices[prices["date"] >= (last - pd.Timedelta(days=days))].copy()
        if len(sub)==0: continue
        sub = sub.set_index("date").resample("W-FRI").last().dropna(subset=["close"]).reset_index()
        if len(sub) < 2: continue
        chg = (sub.iloc[-1]["close"] / sub.iloc[0]["close"] - 1) * 100
        rows.append({"區間":label, "週數":len(sub),
                     "起始":round(sub.iloc[0]["close"],2),
                     "現價":round(sub.iloc[-1]["close"],2),
                     "區間漲跌%":round(chg,1)})
    return pd.DataFrame(rows)


def buy_signal(prices):
    if prices is None or len(prices) < 200: return None
    p = prices["close"].values
    cur = p[-1]; ma50 = p[-50:].mean(); ma200 = p[-200:].mean()
    delta = pd.Series(p).diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = (-delta.clip(upper=0)).rolling(14).mean()
    rsi = (100 - 100/(1 + gain/loss)).iloc[-1]
    p52 = prices[prices["date"] >= prices["date"].max() - pd.Timedelta(days=365)]["close"]
    hi52, lo52 = p52.max(), p52.min()
    sig = []
    if cur < ma50 < ma200: sig.append("🔴 空頭排列")
    elif cur > ma50 > ma200: sig.append("🟢 多頭排列")
    if rsi < 30: sig.append(f"🟢 超賣(RSI {rsi:.0f})")
    elif rsi > 70: sig.append(f"🔴 超買(RSI {rsi:.0f})")
    if cur < lo52 * 1.05: sig.append("🟢 接近52週低")
    if cur > hi52 * 0.97: sig.append("🔴 接近52週高")
    return pd.DataFrame([{
        "現價":round(cur,2),"MA50":round(ma50,2),"MA200":round(ma200,2),
        "距MA50%":round((cur/ma50-1)*100,1),"距MA200%":round((cur/ma200-1)*100,1),
        "RSI14":round(rsi,1),
        "52週高":round(hi52,2),"距52高%":round((cur/hi52-1)*100,1),
        "52週低":round(lo52,2),"距52低%":round((cur/lo52-1)*100,1),
        "綜合訊號":" | ".join(sig) if sig else "中性"
    }])


def analyze(code):
    print(f"\n{'='*60}\n  {code} 深度分析\n{'='*60}")
    rev = annual_revenue_eps(code)
    mon = monthly_revenue(code)
    qtr = latest_quarters(code)
    prices = fetch_prices(code)
    ret = long_term_returns(prices)
    wk = recent_weekly(prices)
    sig = buy_signal(prices)

    if rev is not None: print("\n[10 年年度營收 EPS]"); print(rev.to_string(index=False))
    if mon is not None: print("\n[近 12 月月營收]");    print(mon.to_string(index=False))
    if qtr is not None: print("\n[近 4 季財報]");        print(qtr.to_string(index=False))
    if ret is not None: print("\n[10/5/3/1y 含息年化]"); print(ret.to_string(index=False))
    if wk is not None:  print("\n[近期周線]");           print(wk.to_string(index=False))
    if sig is not None: print("\n[買進訊號]");           print(sig.to_string(index=False))

    out = f"{OUT_DIR}/{code}_深度.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        if rev is not None: rev.to_excel(xw, sheet_name="10年營收EPS", index=False)
        if mon is not None: mon.to_excel(xw, sheet_name="近12月營收", index=False)
        if qtr is not None: qtr.to_excel(xw, sheet_name="近4季財報", index=False)
        if ret is not None: ret.to_excel(xw, sheet_name="長期報酬", index=False)
        if wk is not None:  wk.to_excel(xw, sheet_name="近期走勢", index=False)
        if sig is not None: sig.to_excel(xw, sheet_name="買進訊號", index=False)
    print(f"\n→ 已輸出 {out}")


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN")
    codes = sys.argv[1:]
    if not codes: print("用法: python tw_stock_deepdive.py 2330 3017 ..."); return
    for c in codes: analyze(c)


if __name__ == "__main__":
    main()
