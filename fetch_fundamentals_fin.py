#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股「金融股」基本面比較抓取器  fetch_fundamentals_fin.py
========================================================
金融股(銀行/壽險/金控)的財報結構與工業股完全不同,故另設一套尺。

【刻意「不」計算的指標 — 對金融股無意義】
  毛利率 / 營益率 / 獲利含金量(OCF÷淨利) / 自由現金流
  → 銀行的「營業現金流」是放款、存款、投資部位的增減,不是現金品質,套照妖鏡會得出荒謬結論。

【本版計算的指標 — 金融股該看的】
  EPS、每股淨值(BVPS)、ROE(季 & 近四季)、ROA、負債比(僅供參考,金融股天生高)、
  PER、PBR、殖利率、月營收 月增%/年增%

【FinMind 抓不到、需另尋來源(公開資訊觀測站 MOPS)的銀行品質指標】
  逾放比(NPL)、備抵呆帳覆蓋率、資本適足率(BIS/CAR)、淨利差(NIM)
  → 這幾項才是判斷銀行「資產品質」的核心,本腳本無法提供,請至 MOPS 或各金控法說會補。

pip install finmind pandas openpyxl requests
"""

import os, time
import pandas as pd

# 預設:國泰金、元大金、台新新光金、永豐金、兆豐金、中信金(可自行替換)
TICKERS = ["2882", "2885", "2887", "2890", "2886", "2891"]
START_DATE = "2024-01-01"
TOKEN = os.environ.get("FINMIND_TOKEN", "")
OUTPUT = "台股金融股比較.xlsx"


# ---------- 取數 ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        dl.login_by_token(api_token=TOKEN)
    return dl

def get_per(dl, sid, start):
    try:
        return dl.taiwan_stock_per_pbr(stock_id=sid, start_date=start)
    except Exception:
        import requests
        h = {"Authorization": f"Bearer {TOKEN}"} if TOKEN else {}
        r = requests.get("https://api.finmindtrade.com/api/v4/data",
                         params={"dataset": "TaiwanStockPER", "data_id": sid, "start_date": start},
                         headers=h, timeout=20)
        return pd.DataFrame(r.json().get("data", []))

def fetch_one(dl, sid, start):
    out = {
        "損益表":    dl.taiwan_stock_financial_statement(stock_id=sid, start_date=start),
        "資產負債表": dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=start),
        "月營收":    dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start),
        "PER":      get_per(dl, sid, start),
    }
    time.sleep(1.0)
    return out


# ---------- 工具 ----------
def pivot(df):
    if df is None or df.empty or "type" not in df.columns:
        return pd.DataFrame()
    return df.pivot_table(index="date", columns="type", values="value", aggfunc="first").sort_index()

def pick(piv, *names):
    for n in names:
        if n in piv.columns:
            return piv[n]
    return pd.Series(index=piv.index, dtype="float64")


# ---------- 金融股經營指標(趨勢) ----------
def fin_metrics(raw):
    inc, bal = pivot(raw["損益表"]), pivot(raw["資產負債表"])
    if inc.empty or bal.empty:
        return pd.DataFrame(), None

    # ★ 金融股稅後淨利欄位是 IncomeAfterTax(單數);工業股才是 IncomeAfterTaxes(複數)
    ni  = pick(inc, "IncomeAfterTax", "IncomeAfterTaxes", "ProfitAfterTax",
                    "TotalConsolidatedProfitForThePeriod", "NetIncome")
    eps = pick(inc, "EPS")
    rev = pick(inc, "Revenue", "NetRevenue", "TotalOperatingRevenue")  # 金融股=淨收益,可能不存在

    ta  = pick(bal, "TotalAssets", "Total_Assets")
    tl  = pick(bal, "TotalLiabilities", "Liabilities", "Total_Liabilities")
    eq  = pick(bal, "Equity", "TotalEquity", "EquityAttributableToOwnersOfParent")
    # ★ 每股淨值改用股本反推股數(普通股股本÷面額10),比用 EPS 反推穩(最新季 EPS 常缺值)
    share_cap = pick(bal, "OrdinaryShare", "CommonStock", "ShareCapitalCommonStock")
    shares = share_cap / 10.0

    m = pd.DataFrame(index=bal.index)               # ★ 以資產負債表日期為基準,對齊較穩
    m["稅後淨利(億)"] = (ni.reindex(bal.index) / 1e8).round(1)
    m["EPS"]        = eps.reindex(bal.index).round(2)
    if rev.notna().any():
        m["淨利率%"] = (ni.reindex(bal.index) / rev.reindex(bal.index) * 100).round(2)
    m["股東權益(億)"] = (eq / 1e8).round(0)
    m["每股淨值"]   = (eq / shares).round(2)
    m["ROE%(季)"]   = (ni.reindex(bal.index) / eq * 100).round(2)
    m["ROA%(季)"]   = (ni.reindex(bal.index) / ta * 100).round(3)
    m["負債比%"]    = (tl / ta * 100).round(1)   # 金融股 90%+ 屬常態,僅供參考
    return m, shares


# ---------- 月營收 月增/年增 ----------
def revenue_growth(raw):
    rv = raw["月營收"]
    if rv is None or rv.empty:
        return pd.DataFrame()
    rv = rv.sort_values("date").reset_index(drop=True)
    rv["月增%"] = (rv["revenue"].pct_change() * 100).round(2)
    rv["年增%"] = (rv["revenue"].pct_change(12) * 100).round(2)
    rv["營收(億)"] = (rv["revenue"] / 1e8).round(2)
    return rv[["date", "revenue_year", "revenue_month", "營收(億)", "月增%", "年增%"]]


# ---------- 跨檔比較(金融股口徑) ----------
def comparison_row(sid, raw, perf, rev):
    row = {"代號": sid}
    if not perf.empty:
        last = perf.iloc[-1]
        eps_series = perf["EPS"].dropna()
        ni_series  = perf["稅後淨利(億)"].dropna()
        eq_last    = last.get("股東權益(億)")
        ta_proxy   = None
        row["近四季EPS"]   = round(eps_series.tail(4).sum(), 2) if len(eps_series) else None
        row["每股淨值"]    = last.get("每股淨值")
        row["季ROE%"]     = last.get("ROE%(季)")
        # 近四季 ROE(年化) = 近四季淨利 / 最新股東權益
        if eq_last and len(ni_series) >= 1:
            row["近四季ROE%"] = round(ni_series.tail(4).sum() / eq_last * 100, 2)
        row["季ROA%"]     = last.get("ROA%(季)")
        row["負債比%"]    = last.get("負債比%")
    per = raw["PER"]
    if per is not None and not per.empty:
        p = per.sort_values("date").iloc[-1]
        row["目前PER"] = p.get("PER")
        row["PBR"]     = p.get("PBR")
        row["殖利率%"]  = p.get("dividend_yield")
    if not rev.empty:
        r = rev.iloc[-1]
        row["最新月營收(億)"] = r["營收(億)"]
        row["月增%"] = r["月增%"]
        row["年增%"] = r["年增%"]
    return row


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    compare, per_stock = [], {}
    for sid in TICKERS:
        print(f"抓取 {sid} ...")
        try:
            raw = fetch_one(dl, sid, START_DATE)
            perf, _ = fin_metrics(raw)
            rev = revenue_growth(raw)
            per_stock[sid] = (raw, perf, rev)
            compare.append(comparison_row(sid, raw, perf, rev))
        except Exception as e:
            print(f"  ! {sid} 失敗:{e}")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        if compare:
            pd.DataFrame(compare).set_index("代號").T.to_excel(xw, sheet_name="金融股比較")
        for sid, (raw, perf, rev) in per_stock.items():
            raw["損益表"].to_excel(xw, sheet_name=f"{sid}_損益表"[:31], index=False)
            raw["資產負債表"].to_excel(xw, sheet_name=f"{sid}_資產負債"[:31], index=False)
            if not perf.empty:
                perf.to_excel(xw, sheet_name=f"{sid}_金融指標"[:31])
            if not rev.empty:
                rev.to_excel(xw, sheet_name=f"{sid}_月營收"[:31], index=False)
    print(f"\n已輸出:{OUTPUT}")
    print("⚠ 提醒:逾放比、資本適足率、淨利差(NIM)FinMind 無,請至公開資訊觀測站(MOPS)補。")
    if compare:
        print("\n金融股比較預覽:")
        print(pd.DataFrame(compare).set_index("代號").T.to_string())


if __name__ == "__main__":
    main()
