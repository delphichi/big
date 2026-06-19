# -*- coding: utf-8 -*-
"""
台股精選 20 檔 財報 + 估值 表 (TW Fundamentals & Valuation for 20 picks)
=======================================================================
針對「月營收年增掃描」篩出、最值得深入研究的 20 檔,一次抓最近 ~5 年的三大表 +
每日 PER/PBR/殖利率,算出 5 年趨勢、近四季經營績效與估值,輸出可排序的跨檔比較表,
另含「逐年營收」「逐年EPS」跨檔對照,以及每檔逐季明細分頁。

資料來源:FinMind(taiwan_stock_financial_statement / balance_sheet /
          cash_flows_statement / per_pbr / month_revenue)。一次呼叫即回傳整段歷史,
          故抓 5 年與抓 3 年的 API 次數相同(每檔 5 次)。
輸出   :data/台股精選20_財報估值.xlsx

每檔算出:
  5 年趨勢 → 5年營收CAGR%、5年平均淨利率%、5年平均ROE%(只取季數=4 的完整年)
  近四季   → 毛利率 / 營益率 / 淨利率、近四季EPS、近四季ROE
  財務結構 → 負債比、流動比(最新季)
  現金流照妖鏡 → 獲利含金量(近四季營業現金流 ÷ 近四季淨利)、近四季自由現金流(億)
  估值     → 目前 PER、PBR、殖利率%
  成長     → 最新月營收年增%

★ 大量抓取務必設環境變數 FINMIND_TOKEN(免費約 300 次/hr、設 token 約 600 次/hr);
  20 檔 × 5 dataset = 約 100 次呼叫,免費額度即可。
"""

import os, time
import pandas as pd
import numpy as np

# ---------- 設定 ----------
TOKEN      = os.environ.get("FINMIND_TOKEN", "")
START_DATE = "2020-01-01"                 # 取 ~5 完整年(2021–2025)+ 年增基期
OUTPUT     = "data/台股精選20_財報估值.xlsx"
RATE_SLEEP = 1.0                          # 每檔間隔(降低撞限流機率)
WRITE_DETAIL = True                       # 另輸出每檔逐季經營績效明細分頁

# 精選 20 檔(代號 → 名稱);來自月營收年增掃描的高一致性/長持續名單
PICKS = [
    ("2330", "台積電"),   ("2308", "台達電"),   ("2345", "智邦"),
    ("3017", "奇鋐"),     ("2383", "台光電"),   ("5274", "信驊"),
    ("2059", "川湖"),     ("2368", "金像電"),   ("6223", "旺矽"),
    ("6197", "佳必琪"),   ("4953", "緯軟"),     ("3293", "鈊象"),
    ("2453", "凌群"),     ("3587", "閎康"),     ("2947", "大樹"),
    ("4129", "聯合"),     ("5903", "全家"),     ("2912", "統一超"),
    ("6446", "藥華藥"),   ("3004", "豐達科"),
]


# ---------- FinMind ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try:
            dl.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("token 登入失敗(改用免費額度):", e)
    return dl

def get_per(dl, sid, start):
    """每日 PER/PBR/殖利率。先試 DataLoader,失敗退回原生 REST。"""
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

def decum(s):
    """台股『現金流量表』是 YTD 累計(Q2=半年、Q3=前三季、Q4=全年);
    轉成單季:同一年『本期 − 上一期』,每年首季維持原值。損益表已是單季,不需處理。"""
    if s is None or len(s) == 0:
        return s
    s = s.sort_index()
    out, prev_y, prev_v = {}, None, None
    for d, v in s.items():
        y = str(d)[:4]
        if pd.isna(v):
            out[d] = v
            continue
        out[d] = (v - prev_v) if (y == prev_y and prev_v is not None) else v
        prev_y, prev_v = y, v
    return pd.Series(out)


# ---------- 逐季經營績效 ----------
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

    # 現金流量表為 YTD 累計 → 先轉單季,之後加總近四季才正確
    ocf  = decum(pick(cf, "CashFlowsFromOperatingActivities",
                          "NetCashFlowsFromOperatingActivities",
                          "CashProvidedByOperatingActivities"))
    capex = decum(pick(cf, "PropertyAndPlantAndEquipment",
                          "AcquisitionOfPropertyPlantAndEquipment",
                          "PaymentsToAcquirePropertyPlantAndEquipment"))

    m = pd.DataFrame(index=inc.index)
    m["營收(億)"]   = (rev / 1e8).round(1)
    m["_毛利"]      = gp
    m["_營益"]      = op
    m["_淨利"]      = ni
    m["毛利率%"]    = (gp / rev * 100).round(2)
    m["營益率%"]    = (op / rev * 100).round(2)
    m["淨利率%"]    = (ni / rev * 100).round(2)
    m["EPS"]       = eps.round(2)
    m["_權益"]      = eq
    m["負債比%"]    = (tl / ta * 100).round(1)
    m["流動比%"]    = (ca / cl * 100).round(0)
    m["_OCF"]       = ocf
    m["營業現金流(億)"] = (ocf / 1e8).round(1)
    m["自由現金流(億)"] = ((ocf + capex) / 1e8).round(1)   # capex 在現金流量表多為負值
    return m


def revenue_yoy(raw):
    rv = raw["月營收"]
    if rv is None or rv.empty:
        return None
    rv = rv.sort_values("date").reset_index(drop=True)
    yoy = rv["revenue"].pct_change(12) * 100
    return round(float(yoy.iloc[-1]), 1) if len(yoy) and pd.notna(yoy.iloc[-1]) else None


# ---------- 逐年彙整(供 5 年比較)----------
def yearly(perf):
    """把逐季 perf 彙整成『逐年』:年度營收(億)/EPS 加總、年末權益、淨利、三率年均、季數。"""
    yd = {}
    for d, r in perf.iterrows():
        y = str(d)[:4]
        o = yd.setdefault(y, {"rev": 0.0, "eps": 0.0, "ni": 0.0, "eq": None,
                              "gm": [], "om": [], "nm": [], "qn": 0})
        o["qn"] += 1
        if pd.notna(r["營收(億)"]): o["rev"] += r["營收(億)"]
        if pd.notna(r["EPS"]):     o["eps"] += r["EPS"]
        if pd.notna(r["_淨利"]):    o["ni"]  += r["_淨利"]
        if pd.notna(r["_權益"]):    o["eq"]   = r["_權益"]          # 年內最後一筆 ≈ 年末權益
        for k, col in (("gm", "毛利率%"), ("om", "營益率%"), ("nm", "淨利率%")):
            if pd.notna(r[col]): o[k].append(r[col])
    return yd


# ---------- 跨檔比較(一檔一列)----------
def summary_row(sid, name, raw):
    row = {"代號": sid, "名稱": name}
    rev_year, eps_year = {}, {}                       # 供「逐年」跨檔對照表
    perf = performance(raw)
    if not perf.empty:
        # ── 5 年趨勢(只取季數=4 的完整年)──
        yd = yearly(perf)
        full = sorted(y for y, o in yd.items() if o["qn"] >= 4)[-5:]   # 最近 5 個完整年
        rev_year = {y: round(yd[y]["rev"], 1) for y in full}
        eps_year = {y: round(yd[y]["eps"], 2) for y in full}
        if len(full) >= 2 and yd[full[0]]["rev"] > 0:
            n = len(full) - 1
            row["5年營收CAGR%"] = round(((yd[full[-1]]["rev"] / yd[full[0]]["rev"]) ** (1 / n) - 1) * 100, 1)
        roe_y = [yd[y]["ni"] / yd[y]["eq"] * 100 for y in full if yd[y]["eq"]]
        nm_y  = [float(np.mean(yd[y]["nm"])) for y in full if yd[y]["nm"]]
        if roe_y: row["5年平均ROE%"]   = round(float(np.mean(roe_y)), 1)
        if nm_y:  row["5年平均淨利率%"] = round(float(np.mean(nm_y)), 1)
        last4 = perf.tail(4)
        # 近四季平均三率(平滑單季波動)
        row["毛利率%"]   = round(float(last4["毛利率%"].mean(skipna=True)), 1)
        row["營益率%"]   = round(float(last4["營益率%"].mean(skipna=True)), 1)
        row["淨利率%"]   = round(float(last4["淨利率%"].mean(skipna=True)), 1)
        # 近四季 EPS、ROE(近四季淨利 ÷ 最新季權益)
        eps4 = last4["EPS"].dropna()
        row["近四季EPS"] = round(float(eps4.sum()), 2) if len(eps4) else None
        ni4  = last4["_淨利"].dropna().sum()
        eqL  = perf["_權益"].dropna()
        if len(eqL) and eqL.iloc[-1]:
            row["近四季ROE%"] = round(ni4 / eqL.iloc[-1] * 100, 1)
        # 財務結構(最新季)
        row["負債比%"] = perf["負債比%"].dropna().iloc[-1] if perf["負債比%"].notna().any() else None
        row["流動比%"] = perf["流動比%"].dropna().iloc[-1] if perf["流動比%"].notna().any() else None
        # 現金流照妖鏡(近四季營業現金流 ÷ 近四季淨利)
        ocf4 = last4["_OCF"].dropna().sum()
        if ni4:
            row["獲利含金量"] = round(ocf4 / ni4, 2)
        row["近四季自由現金流(億)"] = round(float(last4["自由現金流(億)"].sum(skipna=True)), 1)
        row["最新季"] = str(perf.index[-1])
    # 估值:目前 PER/PBR/殖利率
    per = raw["PER"]
    if per is not None and not per.empty:
        p = per.sort_values("date").iloc[-1]
        row["PER"]    = p.get("PER")
        row["PBR"]    = p.get("PBR")
        row["殖利率%"] = p.get("dividend_yield")
    # 成長:最新月營收年增
    row["最新月營收年增%"] = revenue_yoy(raw)
    return row, perf, rev_year, eps_year


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    rows, details = [], {}
    rev_years, eps_years = {}, {}                      # {代號名稱: {年: 值}} 供逐年對照
    for i, (sid, name) in enumerate(PICKS, 1):
        print(f"[{i}/{len(PICKS)}] 抓取 {sid} {name} ...")
        for attempt in range(3):
            try:
                raw = fetch_one(dl, sid, START_DATE)
                row, perf, rev_y, eps_y = summary_row(sid, name, raw)
                rows.append(row)
                label = f"{sid} {name}"
                if rev_y: rev_years[label] = rev_y
                if eps_y: eps_years[label] = eps_y
                if WRITE_DETAIL and not perf.empty:
                    details[f"{sid}_{name}"] = perf[[c for c in perf.columns if not c.startswith("_")]]
                break
            except Exception as e:
                msg = str(e).lower()
                if any(k in msg for k in ("limit", "402", "429", "too many", "exceed")) and attempt < 2:
                    print(f"  ! 疑似限流,等 60s 重試({attempt+1}/2):{e}")
                    time.sleep(60); continue
                print(f"  ! {sid} 失敗:{e}")
                rows.append({"代號": sid, "名稱": name})
                break
        time.sleep(RATE_SLEEP)

    df = pd.DataFrame(rows)
    cols = ["代號", "名稱", "最新季",
            "5年營收CAGR%", "5年平均淨利率%", "5年平均ROE%",          # ← 5 年趨勢
            "毛利率%", "營益率%", "淨利率%", "近四季EPS", "近四季ROE%",  # ← 近四季快照
            "負債比%", "流動比%", "獲利含金量", "近四季自由現金流(億)",
            "PER", "PBR", "殖利率%", "最新月營收年增%"]
    df = df[[c for c in cols if c in df.columns]]
    for c in df.columns:
        if c not in ("代號", "名稱", "最新季"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    sort_key = "5年平均ROE%" if "5年平均ROE%" in df.columns else "近四季ROE%"
    df = df.sort_values(sort_key, ascending=False, na_position="last")

    def pivot_years(d):
        if not d:
            return pd.DataFrame()
        years = sorted({y for v in d.values() for y in v})
        out = pd.DataFrame({lbl: pd.Series(v) for lbl, v in d.items()}).T
        return out.reindex(columns=years)

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="財報估值比較", index=False)
        ry, ey = pivot_years(rev_years), pivot_years(eps_years)
        if not ry.empty: ry.to_excel(xw, sheet_name="逐年營收(億)")
        if not ey.empty: ey.to_excel(xw, sheet_name="逐年EPS")
        for key, perf in details.items():
            perf.to_excel(xw, sheet_name=key[:31])
    print(f"\n完成 → {OUTPUT}({len(df)} 檔)")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
