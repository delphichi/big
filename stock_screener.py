#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股四關篩選器  stock_screener.py
=================================
把「持續成長 × 有現金 × 估值不貴(相對自己歷史) × 高品質(相對自己歷史)」四關程式化。

關卡定義(可在 RULES 調整):
  關1 成長:近 12 個月,月營收 YoY 為正的月份數 ≥ 10
  關2 現金:近四季 獲利含金量(ΣOCF/Σ淨利) ≥ 0.8  且  近四季自由現金流 > 0
           (用「近四季」而非單季,避開台股 Q1 季節性偏弱)
  關3 估值:目前 PE 與 PB,各自落在「自己歷史」的百分位 ≤ 50(即不貴於自身中位數)
  關4 品質:近四季 ROE 與 ROIC,各自落在「自己歷史」的百分位 ≥ 50(即優於自身中位數)

四關全過 → ★合格;否則列出卡在哪關。

資料源:FinMind(月營收、每日PER/PBR、損益表、資產負債表、現金流量表)。
歷史長度建議 5 年以上,百分位才有意義。

pip install finmind pandas openpyxl requests
"""

import os, time
import numpy as np
import pandas as pd

TICKERS = ["2408", "2379", "2327", "9942", "6274"]   # 範例:南亞科 瑞昱 國巨 茂順 台燿
START_DATE = "2019-01-01"     # 取 5+ 年,供歷史百分位
TOKEN = os.environ.get("FINMIND_TOKEN", "")
OUTPUT = "台股四關篩選.xlsx"

RULES = dict(grow_months=10, grow_window=12, cashq_min=0.8,
             val_pct_max=50, qual_pct_min=50)


# ---------- 取數 ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        dl.login_by_token(api_token=TOKEN)
    return dl

def get_per(dl, sid, start):
    try:
        df = dl.taiwan_stock_per_pbr(stock_id=sid, start_date=start)
        if df is not None and not df.empty:
            return df
    except Exception:
        pass
    import requests
    h = {"Authorization": f"Bearer {TOKEN}"} if TOKEN else {}
    r = requests.get("https://api.finmindtrade.com/api/v4/data",
                     params={"dataset": "TaiwanStockPER", "data_id": sid, "start_date": start},
                     headers=h, timeout=30)
    return pd.DataFrame(r.json().get("data", []))

def fetch_one(dl, sid, start):
    out = {
        "inc": dl.taiwan_stock_financial_statement(stock_id=sid, start_date=start),
        "bal": dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=start),
        "cf":  dl.taiwan_stock_cash_flows_statement(stock_id=sid, start_date=start),
        "rev": dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start),
        "per": get_per(dl, sid, start),
    }
    time.sleep(1.2)
    return out


# ---------- 工具 ----------
def pivot(df):
    if df is None or df.empty or "type" not in df.columns:
        return pd.DataFrame()
    return df.pivot_table(index="date", columns="type", values="value", aggfunc="first").sort_index()

def pick(p, *names):
    for n in names:
        if n in p.columns:
            return p[n]
    return pd.Series(index=p.index, dtype="float64")

def pctile(series, value):
    """value 落在歷史 series 的百分位(0-100);越低=越接近歷史低點。"""
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 8 or value is None or pd.isna(value):
        return None
    return round((s <= value).mean() * 100, 1)


# ---------- 關1:營收動能 ----------
def gate_growth(raw):
    rv = raw["rev"]
    if rv is None or rv.empty:
        return None, None
    rv = rv.sort_values("date")
    yoy = rv["revenue"].pct_change(12) * 100
    last = yoy.dropna().tail(RULES["grow_window"])
    pos = int((last > 0).sum())
    return pos, len(last)


# ---------- 關2:現金(近四季) ----------
def gate_cash(raw):
    inc, cf = pivot(raw["inc"]), pivot(raw["cf"])
    if inc.empty or cf.empty:
        return None, None
    ni  = pick(inc, "IncomeAfterTaxes", "IncomeAfterTax", "ProfitAfterTax")
    ocf = pick(cf, "CashFlowsFromOperatingActivities",
                   "NetCashFlowsFromOperatingActivities", "CashProvidedByOperatingActivities")
    cap = pick(cf, "PropertyAndPlantAndEquipment",
                   "AcquisitionOfPropertyPlantAndEquipment", "PaymentsToAcquirePropertyPlantAndEquipment")
    ni4  = ni.dropna().tail(4).sum()
    ocf4 = ocf.dropna().tail(4).sum()
    cap4 = cap.dropna().tail(4).sum()
    cashq = round(ocf4 / ni4, 2) if ni4 else None
    fcf4  = round((ocf4 + cap4) / 1e8, 1)          # capex 多為負值
    return cashq, fcf4


# ---------- 關3 & 4:估值/品質 + 歷史百分位 ----------
def roe_roic_series(raw):
    """回傳歷史的 近四季ROE / ROA / ROIC 序列(季度索引),供百分位用。"""
    inc, bal = pivot(raw["inc"]), pivot(raw["bal"])
    if inc.empty or bal.empty:
        return pd.DataFrame()
    ni  = pick(inc, "IncomeAfterTaxes", "IncomeAfterTax", "ProfitAfterTax")
    op  = pick(inc, "OperatingIncome")
    pre = pick(inc, "PreTaxIncome", "IncomeBeforeTax")
    eq  = pick(bal, "Equity", "TotalEquity")
    ta  = pick(bal, "TotalAssets")
    cl  = pick(bal, "CurrentLiabilities")
    cash = pick(bal, "CashAndCashEquivalents")

    idx = bal.index
    ni4 = ni.reindex(inc.index).rolling(4).sum().reindex(idx)
    op4 = op.reindex(inc.index).rolling(4).sum().reindex(idx)
    pre4 = pre.reindex(inc.index).rolling(4).sum().reindex(idx)
    taxrate = (1 - (ni4 / pre4)).clip(0, 0.4)       # 有效稅率,限縮在合理區間
    nopat = op4 * (1 - taxrate)
    invested = (ta - cl - cash)                     # 投入資本近似 = 總資產 − 流動負債 − 現金

    out = pd.DataFrame(index=idx)
    out["ROE"]  = (ni4 / eq * 100)
    out["ROA"]  = (ni4 / ta * 100)
    out["ROIC"] = (nopat / invested * 100)
    return out.dropna(how="all")


def screen_one(sid, raw):
    r = {"代號": sid}
    # 關1
    pos, win = gate_growth(raw)
    r["營收正成長月數"] = f"{pos}/{win}" if pos is not None else "—"
    g1 = (pos is not None and pos >= RULES["grow_months"])
    # 關2
    cashq, fcf = gate_cash(raw)
    r["獲利含金量"] = cashq; r["近四季FCF(億)"] = fcf
    g2 = (cashq is not None and cashq >= RULES["cashq_min"] and fcf is not None and fcf > 0)
    # 關3:估值 + 百分位
    per = raw["per"]
    pe = pb = pe_p = pb_p = None
    if per is not None and not per.empty:
        per = per.sort_values("date")
        pe = per["PER"].iloc[-1]; pb = per["PBR"].iloc[-1]
        pe_p = pctile(per["PER"], pe); pb_p = pctile(per["PBR"], pb)
    r["PE"] = round(pe, 1) if pe else None; r["PE歷史百分位"] = pe_p
    r["PB"] = round(pb, 2) if pb else None; r["PB歷史百分位"] = pb_p
    g3 = (pe_p is not None and pb_p is not None and
          pe_p <= RULES["val_pct_max"] and pb_p <= RULES["val_pct_max"])
    # 關4:品質 + 百分位
    q = roe_roic_series(raw)
    roe = roic = roa = roe_p = roic_p = None
    if not q.empty:
        roe = q["ROE"].dropna().iloc[-1] if q["ROE"].notna().any() else None
        roa = q["ROA"].dropna().iloc[-1] if q["ROA"].notna().any() else None
        roic = q["ROIC"].dropna().iloc[-1] if q["ROIC"].notna().any() else None
        roe_p = pctile(q["ROE"], roe); roic_p = pctile(q["ROIC"], roic)
    r["近四季ROE%"] = round(roe, 1) if roe is not None else None
    r["近四季ROA%"] = round(roa, 2) if roa is not None else None
    r["近四季ROIC%"] = round(roic, 1) if roic is not None else None
    r["ROE歷史百分位"] = roe_p; r["ROIC歷史百分位"] = roic_p
    g4 = (roe_p is not None and roic_p is not None and
          roe_p >= RULES["qual_pct_min"] and roic_p >= RULES["qual_pct_min"])
    # 綜合
    gates = {"成長": g1, "現金": g2, "估值": g3, "品質": g4}
    passed = [k for k, v in gates.items() if v]
    failed = [k for k, v in gates.items() if not v]
    r["通過關卡"] = "／".join(passed) if passed else "無"
    r["卡在"] = "／".join(failed) if failed else ""
    r["★四關全過"] = "★ 是" if not failed else "否"
    return r


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    rows = []
    for sid in TICKERS:
        print(f"篩選 {sid} ...")
        try:
            rows.append(screen_one(sid, fetch_one(dl, sid, START_DATE)))
        except Exception as e:
            print(f"  ! {sid} 失敗:{e}")
    df = pd.DataFrame(rows)
    cols = ["代號", "★四關全過", "通過關卡", "卡在", "營收正成長月數", "獲利含金量",
            "近四季FCF(億)", "PE", "PE歷史百分位", "PB", "PB歷史百分位",
            "近四季ROE%", "ROE歷史百分位", "近四季ROIC%", "ROIC歷史百分位", "近四季ROA%"]
    df = df[[c for c in cols if c in df.columns]]
    df.to_excel(OUTPUT, sheet_name="四關篩選", index=False)
    print(f"\n已輸出:{OUTPUT}\n")
    pd.set_option("display.unicode.east_asian_width", True); pd.set_option("display.width", 220)
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
