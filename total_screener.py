#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股五維總篩選  total_screener.py
=================================
一張表整合「基本面四關」+「資金面共振」,並自動分類標的性質。

  基本面(來自財報/月營收/PER)
    關1 成長:近12月,月營收YoY為正 ≥10個月
    關2 現金:近四季 獲利含金量(ΣOCF/Σ淨利)≥0.8 且 近四季自由現金流>0
    關3 估值:目前PE、PB 各自落在「自己歷史」百分位 ≤50(不貴於自身中位數)
    關4 品質:近四季 ROE、ROIC 各自落在「自己歷史」百分位 ≥50(優於自身中位數)
  資金面(來自每日股價,不需財報)
    關5 共振:RS×週斜率×均線×量能×大盤 共10條件,共振分數 ≥ 門檻

  總評分 = 基本面通過比 ×50 + 共振分數 ×0.5   (0-100)

  自動分類:
    ★ 主流   = 基本面≥3關 且 資金面達標   → 基本面與資金同向(最強)
    ⚠ 純資金 = 資金面達標 但 基本面≤1關    → 投機/軋空,慎防假突破
    ◎ 潛伏   = 基本面≥3關 但 資金面未達標  → 好公司資金未到,觀察清單
    ○ 偏多 / — 不符

資料源:FinMind。大盤基準預設 0050。歷史建議 5 年(百分位)+ 2 年股價(斜率/52週高)。
pip install finmind pandas numpy openpyxl requests
"""

import os, time
import numpy as np
import pandas as pd

TICKERS = ["2412", "2912", "1476", "1560", "9942", "2360"]  # 台積電 聯發科 瑞昱 南亞科 緯創 台燿
START_FUND = "2019-01-01"     # 財報/月營收/PER 歷史(算百分位)
START_PRICE = "2023-06-01"    # 每日股價(算斜率/52週高)
BENCHMARK = "0050"
TOKEN = os.environ.get("FINMIND_TOKEN", "")
OUTPUT = "data/台股五維總篩選.xlsx"      # 直接輸出到 data/ 資料夾

RULES = dict(grow_months=10, grow_window=12, cashq_min=0.8,
             val_pct_max=50, qual_pct_min=50, resonance_min=70,
             # 9 類分類用門檻
             val_pct_extreme=85,    # PE 歷史百分位 ≥ 此值 = 估值極高(透支)
             reson_cold=50,         # 共振 < 此值 = 資金冷
             low_growth_months=6,   # 近12月正成長 ≤ 此值 = 成長弱(收租型)
             yield_min=4.0)         # 殖利率 ≥ 此值(%) = 高息(收租型)
D_YEAR, D_13W, D_26W = 252, 65, 130


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

def fetch_price(dl, sid, start):
    df = dl.taiwan_stock_daily(stock_id=sid, start_date=start)
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy(); df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").set_index("date")
    vol = "Trading_Volume" if "Trading_Volume" in df.columns else "volume"
    out = pd.DataFrame({"close": df["close"], "vol": df[vol]})
    return out[out["close"] > 0]

def fetch_all(dl, sid):
    out = {
        "inc": dl.taiwan_stock_financial_statement(stock_id=sid, start_date=START_FUND),
        "bal": dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=START_FUND),
        "cf":  dl.taiwan_stock_cash_flows_statement(stock_id=sid, start_date=START_FUND),
        "rev": dl.taiwan_stock_month_revenue(stock_id=sid, start_date=START_FUND),
        "per": get_per(dl, sid, START_FUND),
        "price": fetch_price(dl, sid, START_PRICE),
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
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 8 or value is None or pd.isna(value):
        return None
    return round((s <= value).mean() * 100, 1)

def log_slope(prices):
    p = pd.to_numeric(prices, errors="coerce").dropna()
    if len(p) < 5:
        return None
    x = np.arange(len(p))
    return float(np.polyfit(x, np.log(p.values), 1)[0])

def to_single_q(s):
    """現金流量表是『年初至今累計』,還原成單季:同年內=本季累計−上季累計,Q1不變。
    (損益表 FinMind 已是單季,不可套用此函式)"""
    s = pd.to_numeric(s, errors="coerce").dropna()
    if s.empty:
        return s
    s.index = pd.to_datetime(s.index); s = s.sort_index()
    out, prev, prevy = [], None, None
    for dt, v in s.items():
        q = (dt.month - 1) // 3 + 1
        out.append(v if (q == 1 or prev is None or dt.year != prevy) else v - prev)
        prev, prevy = v, dt.year
    return pd.Series(out, index=s.index)

def find_capex(cf):
    """穩健抓資本支出:先試精確名,再模糊比對(含 Propert + Plant/Equipment),
    取現金流出最大(總和最負)那條,避免漏抓或誤抓處分利益。"""
    for n in ("AcquisitionOfPropertyPlantAndEquipment",
              "PaymentsToAcquirePropertyPlantAndEquipment",
              "PurchaseOfPropertyPlantAndEquipment",
              "PropertyAndPlantAndEquipment"):
        if n in cf.columns:
            return cf[n]
    cands = [c for c in cf.columns
             if "Propert" in c and ("Plant" in c or "Equipment" in c) and not c.endswith("_per")]
    if cands:
        best = min(cands, key=lambda c: pd.to_numeric(cf[c], errors="coerce").sum())  # 最負=最大流出
        return cf[best]
    return pd.Series(index=cf.index, dtype="float64")


# ---------- 基本面四關 ----------
def gate_growth(raw):
    rv = raw["rev"]
    if rv is None or rv.empty:
        return None, None
    rv = rv.sort_values("date")
    yoy = rv["revenue"].pct_change(12) * 100
    last = yoy.dropna().tail(RULES["grow_window"])
    return int((last > 0).sum()), len(last)

def gate_cash(raw):
    inc, cf = pivot(raw["inc"]), pivot(raw["cf"])
    if inc.empty or cf.empty:
        return None, None, {}
    ni  = pick(inc, "IncomeAfterTaxes", "IncomeAfterTax", "ProfitAfterTax")
    ocf = pick(cf, "CashFlowsFromOperatingActivities",
                   "NetCashFlowsFromOperatingActivities", "CashProvidedByOperatingActivities")
    cap = find_capex(cf)
    # 現金流量表是累計值 → 先還原單季再加總;損益表已是單季,不動
    ni4  = ni.dropna().tail(4).sum()
    ocf4 = to_single_q(ocf).tail(4).sum()
    cap4 = to_single_q(cap).tail(4).sum()
    cashq = round(ocf4 / ni4, 2) if ni4 else None
    fcf4  = round((ocf4 + cap4) / 1e8, 1)
    audit = {"近四季OCF(億)": round(ocf4/1e8,1), "近四季淨利(億)": round(ni4/1e8,1),
             "近四季capex(億)": round(cap4/1e8,1), "capex欄位": cap.name}
    return cashq, fcf4, audit

def roe_roic_series(raw):
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
    taxrate = (1 - (ni4 / pre4)).clip(0, 0.4)
    nopat = op4 * (1 - taxrate)
    invested = (ta - cl - cash)
    out = pd.DataFrame(index=idx)
    out["ROE"]  = (ni4 / eq * 100)
    out["ROA"]  = (ni4 / ta * 100)
    out["ROIC"] = (nopat / invested * 100)
    return out.dropna(how="all")


# ---------- 資金面共振 ----------
def momentum(raw, bench):
    stock = raw["price"]
    if stock is None or stock.empty or len(stock) < D_26W:
        return None
    s, v = stock["close"], stock["vol"]
    o = {}
    aligned = pd.concat([s.rename("s"), bench.rename("b")], axis=1).dropna()
    rs = aligned["s"] / aligned["b"]
    o["近半年相對報酬%"] = None
    if len(aligned) > D_26W:
        o["近半年相對報酬%"] = round((aligned["s"].iloc[-1]/aligned["s"].iloc[-D_26W]
                                  - aligned["b"].iloc[-1]/aligned["b"].iloc[-D_26W]) * 100, 1)
    rs_nh = (len(rs) >= D_YEAR and rs.iloc[-1] >= rs.tail(D_YEAR).max() * 0.98)
    rs_up = (len(rs) > D_13W and rs.iloc[-1] > rs.iloc[-D_13W])
    wk = s.resample("W-FRI").last().dropna()
    sl_r = log_slope(wk.tail(13)); sl_p = log_slope(wk.iloc[-26:-13]) if len(wk) >= 26 else None
    o["週斜率%/週"] = round(sl_r*100, 2) if sl_r is not None else None
    accel = (sl_r is not None and sl_p is not None and sl_r > sl_p)
    wk_pos = (sl_r is not None and sl_r > 0)
    d_pos = (log_slope(s.tail(20)) or 0) > 0
    ma50, ma200, price = s.tail(50).mean(), s.tail(200).mean(), s.iloc[-1]
    p_nh = (len(s) >= D_YEAR and price >= s.tail(D_YEAR).max() * 0.95)
    vol_surge = v.tail(20).mean() > v.tail(60).mean()
    bench_up = (len(bench) >= 200 and bench.iloc[-1] > bench.tail(200).mean())
    conds = {"站上200日線": price > ma200, "站上50日線": price > ma50, "均線多頭": ma50 > ma200,
             "價格創52週高": p_nh, "RS創52週高": rs_nh, "RS上升": rs_up,
             "週斜率為正": wk_pos, "週斜率加速": accel, "量能放大": vol_surge, "大盤多頭": bench_up}
    o["共振分數"] = round(sum(bool(x) for x in conds.values())/len(conds)*100)
    o["RS創高"] = "✔" if rs_nh else ""
    o["RS上升"] = "✔" if rs_up else ""
    o["亮燈"] = "／".join(k for k, x in conds.items() if x)
    return o


# ---------- 組裝一檔 ----------
def assess(sid, raw, bench):
    r = {"代號": sid}
    # 基本面
    pos, win = gate_growth(raw)
    r["營收正成長"] = f"{pos}/{win}" if pos is not None else "—"
    g1 = pos is not None and pos >= RULES["grow_months"]
    cashq, fcf, cash_audit = gate_cash(raw)
    r["含金量"], r["近四季FCF"] = cashq, fcf
    g2 = cashq is not None and cashq >= RULES["cashq_min"] and fcf is not None and fcf > 0
    per = raw["per"]; pe = pb = pe_p = pb_p = divy = None
    if per is not None and not per.empty:
        per = per.sort_values("date"); pe = per["PER"].iloc[-1]; pb = per["PBR"].iloc[-1]
        pe_p, pb_p = pctile(per["PER"], pe), pctile(per["PBR"], pb)
        if "dividend_yield" in per.columns:
            divy = per["dividend_yield"].iloc[-1]
    r["PE"], r["PE百分位"] = (round(pe,1) if pe else None), pe_p
    r["PB百分位"] = pb_p
    r["殖利率%"] = round(divy,2) if divy is not None and not pd.isna(divy) else None
    g3 = pe_p is not None and pb_p is not None and pe_p <= RULES["val_pct_max"] and pb_p <= RULES["val_pct_max"]
    q = roe_roic_series(raw); roe = roic = roe_p = roic_p = None
    if not q.empty:
        if q["ROE"].notna().any():  roe = q["ROE"].dropna().iloc[-1]
        if q["ROIC"].notna().any(): roic = q["ROIC"].dropna().iloc[-1]
        roe_p, roic_p = pctile(q["ROE"], roe), pctile(q["ROIC"], roic)
    r["ROE%"], r["ROE百分位"] = (round(roe,1) if roe is not None else None), roe_p
    r["ROIC%"], r["ROIC百分位"] = (round(roic,1) if roic is not None else None), roic_p
    g4 = roe_p is not None and roic_p is not None and roe_p >= RULES["qual_pct_min"] and roic_p >= RULES["qual_pct_min"]

    # 資金面
    m = momentum(raw, bench)
    reson = m["共振分數"] if m else None
    if m:
        r["共振分數"] = reson; r["相對報酬%"] = m["近半年相對報酬%"]
        r["週斜率%"] = m["週斜率%/週"]; r["RS創高"] = m["RS創高"]; r["亮燈"] = m["亮燈"]
    g5 = reson is not None and reson >= RULES["resonance_min"]

    # 綜合
    fund = sum([g1, g2, g3, g4])
    gates = {"成長": g1, "現金": g2, "估值": g3, "品質": g4, "資金": g5}
    r["基本面"] = f"{fund}/4"
    r["五關"] = "".join(("✔" if v else "·") for v in [g1, g2, g3, g4, g5]) + "  (成長現金估值品質資金)"
    r["卡在"] = "／".join(k for k, v in gates.items() if not v) or "全過"
    score = round(fund/4*50 + (reson/100*50 if reson is not None else 0))
    r["總評分"] = score

    # ---- 9 類分類(依優先序,先中先定) ----
    burn_now   = (cashq is not None and cashq < RULES["cashq_min"]) and (fcf is not None and fcf <= 0)
    cheap      = g3
    val_extreme = (pe_p is not None and pe_p >= RULES["val_pct_extreme"])
    hot        = g5
    cold       = (reson is not None and reson < RULES["reson_cold"])
    growing    = g1
    low_growth = (pos is not None and pos <= RULES["low_growth_months"])
    high_yield = (divy is not None and not pd.isna(divy) and divy >= RULES["yield_min"])
    qual_ok    = g4

    if growing and burn_now:
        cat = "🔥 燒錢成長(成長吞現金)"
    elif hot and fund <= 1:
        cat = "⚠ 純資金(投機/慎防假突破)"
    elif cheap and not qual_ok and not growing:
        cat = "🪤 價值陷阱(便宜有原因)"
    elif fund >= 2 and val_extreme and hot:
        cat = "🏔 估值透支(好公司但太貴)"
    elif fund >= 3 and hot:
        cat = "★ 主流(基本面+資金同向)"
    elif fund >= 3 and cheap and not hot:
        cat = "◎ 潛伏(好公司便宜·資金未到)"
    elif growing and cheap and not qual_ok:
        cat = "🔄 循環反轉初期(便宜·營收回溫·獲利未復)"
    elif qual_ok and g2 and cheap and not hot:
        cat = "💎 價值低估(績優便宜·被冷落)"
    elif low_growth and g2 and high_yield and not hot:
        cat = "🐢 穩定收租(牛皮績優·股息)"
    elif fund >= 2:
        cat = "○ 偏多"
    else:
        cat = "— 不符"
    r["分類"] = cat

    # 計算稽核(中間值,供逐項核對)
    def last_date(df, col="date"):
        try:    return str(pd.to_datetime(df[col]).max().date())
        except Exception: return "—"
    per_df = raw["per"]
    audit = {"代號": sid,
             **cash_audit,
             "→含金量(OCF/淨利)": cashq, "→近四季FCF(億)": fcf,
             "PE現值": r.get("PE"), "PE歷史百分位": r.get("PE百分位"), "PB歷史百分位": r.get("PB百分位"),
             "近四季ROE%": r.get("ROE%"), "近四季ROIC%": r.get("ROIC%")}
    fresh = {"代號": sid,
             "財報最新季": last_date(raw["inc"]) if raw["inc"] is not None and not raw["inc"].empty else "—",
             "月營收最新": last_date(raw["rev"]) if raw["rev"] is not None and not raw["rev"].empty else "—",
             "股價最新": str(raw["price"].index.max().date()) if raw["price"] is not None and not raw["price"].empty else "—",
             "PER最新": last_date(per_df) if per_df is not None and not per_df.empty else "—"}
    return r, audit, fresh


# ---------- 主流程 ----------
DUMP_RAW = True          # 是否輸出每檔原始三表(供逐列驗證);檔案會變大
RAW_TAIL = 12            # 原始表只保留最近 N 季,控制檔案大小

def main():
    dl = make_loader()
    bench = fetch_price(dl, BENCHMARK, START_PRICE)
    bench = bench["close"] if not bench.empty else pd.Series(dtype=float)
    rows, audits, freshes, raws = [], [], [], {}
    for sid in TICKERS:
        print(f"評估 {sid} ...")
        try:
            raw = fetch_all(dl, sid)
            r, audit, fresh = assess(sid, raw, bench)
            rows.append(r); audits.append(audit); freshes.append(fresh); raws[sid] = raw
        except Exception as e:
            print(f"  ! {sid} 失敗:{e}")

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("總評分", ascending=False)
        cols = ["代號", "總評分", "分類", "五關", "卡在", "基本面", "共振分數",
                "營收正成長", "含金量", "近四季FCF", "PE", "PE百分位", "PB百分位",
                "ROE%", "ROE百分位", "ROIC%", "ROIC百分位", "殖利率%",
                "相對報酬%", "週斜率%", "RS創高", "亮燈"]
        df = df[[c for c in cols if c in df.columns]]

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)   # 確保 data/ 存在
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="五維總篩選", index=False)
        # 計算稽核:中間值,可自行除一遍對
        if audits:
            acols = ["代號", "近四季OCF(億)", "近四季淨利(億)", "近四季capex(億)", "capex欄位",
                     "→含金量(OCF/淨利)", "→近四季FCF(億)", "PE現值", "PE歷史百分位",
                     "PB歷史百分位", "近四季ROE%", "近四季ROIC%"]
            ad = pd.DataFrame(audits)
            ad[[c for c in acols if c in ad.columns]].to_excel(xw, sheet_name="計算稽核", index=False)
        # 資料時效:看各資料集最新到哪天(滯後性一眼可見)
        if freshes:
            pd.DataFrame(freshes).to_excel(xw, sheet_name="資料時效", index=False)
        # 每檔原始三表 + 月營收(逐列核對來源)
        if DUMP_RAW:
            for sid, raw in raws.items():
                for key, label in [("inc","損益表"),("bal","資產負債"),("cf","現金流"),("rev","月營收")]:
                    d = raw.get(key)
                    if d is not None and not d.empty:
                        d.tail(RAW_TAIL*40 if key!="rev" else 30).to_excel(
                            xw, sheet_name=f"{sid}_{label}"[:31], index=False)
    print(f"\n已輸出:{OUTPUT}\n")
    pd.set_option("display.unicode.east_asian_width", True); pd.set_option("display.width", 260)
    show = [c for c in ["代號","總評分","分類","五關","卡在","基本面","共振分數"] if c in df.columns]
    print(df[show].to_string(index=False))


if __name__ == "__main__":
    main()
