#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股基本面「全表 + 跨檔比較」抓取器 v2
=====================================
直接抓：損益表、資產負債表、現金流量表、月營收、每日PER/PBR/殖利率
自己算：經營績效比率(毛利率/營益率/淨利率/ROE/負債比/流動比) +
        現金流照妖鏡(營業現金流÷淨利、自由現金流) + 月增%/年增% + 近四季EPS

輸出 Excel：
  每檔 → 損益表 / 資產負債表 / 現金流量表 / 月營收 / 經營績效(趨勢)
  另加 → 「跨檔比較」一張總表，所有股票並排比最新一季的關鍵指標

pip install finmind pandas openpyxl requests
"""

import os, time
import pandas as pd

TICKERS = ["6173", "2344", "3090", "2408", "6770", "3231"]   # 群創 國巨 緯穎 瑞昱（可自行替換）
START_DATE = "2024-01-01"
TOKEN = os.environ.get("FINMIND_TOKEN", "")   # 建議設環境變數；空字串也能跑(額度低)
OUTPUT = "台股基本面_v2.xlsx"


# ---------- 取數 ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        dl.login_by_token(api_token=TOKEN)
    return dl

def get_per(dl, sid, start):
    """每日PER/PBR/殖利率。先試 DataLoader 方法，失敗退回原生 REST。"""
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
        "現金流量表": dl.taiwan_stock_cash_flows_statement(stock_id=sid, start_date=start),
        "月營收":    dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start),
        "PER":      get_per(dl, sid, start),
    }
    time.sleep(1.0)
    return out


# ---------- 工具：長表轉寬表 + 容錯取欄 ----------
def pivot(df):
    if df is None or df.empty or "type" not in df.columns:
        return pd.DataFrame()
    return df.pivot_table(index="date", columns="type", values="value", aggfunc="first").sort_index()

def pick(piv, *names):
    """回傳第一個存在的欄位 Series；都不存在則回傳全 NaN。"""
    for n in names:
        if n in piv.columns:
            return piv[n]
    return pd.Series(index=piv.index, dtype="float64")


# ---------- 經營績效 + 現金流照妖鏡 ----------
def performance(raw):
    inc, bal, cf = pivot(raw["損益表"]), pivot(raw["資產負債表"]), pivot(raw["現金流量表"])
    if inc.empty:
        return pd.DataFrame()

    rev  = pick(inc, "Revenue")
    gp   = pick(inc, "GrossProfit")
    op   = pick(inc, "OperatingIncome")
    ni   = pick(inc, "IncomeAfterTaxes", "ProfitAfterTax", "NetIncome")
    eps  = pick(inc, "EPS")

    ta   = pick(bal, "TotalAssets", "Total_Assets")
    tl   = pick(bal, "TotalLiabilities", "Liabilities", "Total_Liabilities")
    eq   = pick(bal, "Equity", "TotalEquity", "EquityAttributableToOwnersOfParent")
    ca   = pick(bal, "CurrentAssets")
    cl   = pick(bal, "CurrentLiabilities")

    ocf  = pick(cf, "CashFlowsFromOperatingActivities",
                    "NetCashFlowsFromOperatingActivities",
                    "CashProvidedByOperatingActivities")
    capex = pick(cf, "PropertyAndPlantAndEquipment",
                     "AcquisitionOfPropertyPlantAndEquipment",
                     "PaymentsToAcquirePropertyPlantAndEquipment")

    m = pd.DataFrame(index=inc.index)
    m["營收(億)"]   = (rev / 1e8).round(1)
    m["毛利率%"]    = (gp / rev * 100).round(2)
    m["營益率%"]    = (op / rev * 100).round(2)
    m["淨利率%"]    = (ni / rev * 100).round(2)
    m["EPS"]       = eps.round(2)
    m["ROE%(季)"]   = (ni / eq * 100).round(2)
    m["負債比%"]    = (tl / ta * 100).round(1)
    m["流動比%"]    = (ca / cl * 100).round(0)
    # ── 現金流照妖鏡 ──
    m["營業現金流(億)"] = (ocf / 1e8).round(1)
    m["獲利含金量(OCF/淨利)"] = (ocf / ni).round(2)      # <1 代表賺帳面、收不到現金
    m["自由現金流(億)"] = ((ocf + capex) / 1e8).round(1)  # capex 在現金流量表多為負值
    return m


def revenue_growth(raw):
    rv = raw["月營收"]
    if rv is None or rv.empty:
        return pd.DataFrame()
    rv = rv.sort_values("date").reset_index(drop=True)
    rv["月增%"] = (rv["revenue"].pct_change() * 100).round(2)
    rv["年增%"] = (rv["revenue"].pct_change(12) * 100).round(2)
    rv["營收(億)"] = (rv["revenue"] / 1e8).round(2)
    return rv[["date", "revenue_year", "revenue_month", "營收(億)", "月增%", "年增%"]]


# ---------- 跨檔比較總表 ----------
def comparison_row(sid, raw, perf, rev):
    row = {"代號": sid}
    if not perf.empty:
        last = perf.iloc[-1]
        for c in ["毛利率%", "營益率%", "淨利率%", "ROE%(季)", "負債比%", "流動比%",
                  "營業現金流(億)", "獲利含金量(OCF/淨利)", "自由現金流(億)"]:
            row[c] = last.get(c)
        # 近四季 EPS
        eps = perf["EPS"].dropna()
        row["近四季EPS"] = round(eps.tail(4).sum(), 2) if len(eps) >= 1 else None
    # 目前 PER / PBR
    per = raw["PER"]
    if per is not None and not per.empty:
        p = per.sort_values("date").iloc[-1]
        row["目前PER"] = p.get("PER")
        row["PBR"] = p.get("PBR")
        row["殖利率%"] = p.get("dividend_yield")
    # 最新月營收 月增/年增
    if not rev.empty:
        r = rev.iloc[-1]
        row["最新月營收(億)"] = r["營收(億)"]
        row["月增%"] = r["月增%"]
        row["年增%"] = r["年增%"]
    return row


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    compare = []
    per_stock = {}
    for sid in TICKERS:
        print(f"抓取 {sid} ...")
        try:
            raw = fetch_one(dl, sid, START_DATE)
            perf = performance(raw)
            rev = revenue_growth(raw)
            per_stock[sid] = (raw, perf, rev)
            compare.append(comparison_row(sid, raw, perf, rev))
        except Exception as e:
            print(f"  ! {sid} 失敗：{e}")

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        # 跨檔比較放第一張
        if compare:
            pd.DataFrame(compare).set_index("代號").T.to_excel(xw, sheet_name="跨檔比較")
        for sid, (raw, perf, rev) in per_stock.items():
            raw["損益表"].to_excel(xw, sheet_name=f"{sid}_損益表"[:31], index=False)
            raw["資產負債表"].to_excel(xw, sheet_name=f"{sid}_資產負債"[:31], index=False)
            raw["現金流量表"].to_excel(xw, sheet_name=f"{sid}_現金流"[:31], index=False)
            if not perf.empty:
                perf.to_excel(xw, sheet_name=f"{sid}_經營績效"[:31])
            if not rev.empty:
                rev.to_excel(xw, sheet_name=f"{sid}_月營收"[:31], index=False)
    print(f"\n已輸出：{OUTPUT}")
    if compare:
        print("\n跨檔比較預覽：")
        print(pd.DataFrame(compare).set_index("代號").T.to_string())


if __name__ == "__main__":
    main()
