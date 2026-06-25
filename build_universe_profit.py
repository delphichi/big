#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全市場獲利初篩  build_universe_profit.py
=====================================================================
用 FinMind date-bulk 端點(一次抓「某季全台股」,不必逐檔),掃出「獲利型」候選,
補「產業白名單(build_universe)」的盲區 —— 例:帆宣(廠務工程)不在能力圈關鍵字,
但 ROE/EPS 紮實,過去從沒被體檢過。獲利初篩用財務體質撈,不用產業/營收成長。

方法(便宜,僅 ~5 次 bulk 呼叫掃完 2000 檔):
  ROE ≈ 近四季淨利(4 季 IncomeAfterTaxes 加總) / 最新季權益(Equity) × 100
  EPS新高 ≈ 近四季EPS 為近 N 季最高(獲利動能向上)
篩選門檻(任一):
  ROE ≥ 15            → 高資本效率
  近四季EPS 創近12季高 → 獲利突破(轉機/成長未被產業標籤抓到)

輸出 data/獲利型候選.txt(代號 + 名稱 + ROE + 是否EPS新高),並標出「不在現有 PICKS」的遺珠。
這些遺珠之後可加進 fetch_fundamentals_tw 的 PICKS 做深度體檢。

★ 需 FINMIND_TOKEN(bulk 端點吃額度);無 token 也能跑但易撞限流。
"""
import os
import sys
from datetime import date, datetime
import requests
import pandas as pd
import numpy as np

TOKEN = os.environ.get("FINMIND_TOKEN", "")
API = "https://api.finmindtrade.com/api/v4/data"
OUT = "data/獲利型候選.txt"
ROE_MIN = 15.0          # 近四季 ROE 門檻
EPS_HIGH_LOOKBACK = 12  # 近四季EPS 是否為近 12 季最高


def recent_quarter_ends(n=13):
    """回傳最近 n 個『已公布』的季末日(字串),最新在前。
    報表落後:Q1~5/15、Q2~8/14、Q3~11/14、Q4~隔年3/31。保守取季末+50天已過者。"""
    today = date.today()
    ends = []
    y = today.year
    for yr in range(y, y - 5, -1):
        for m, d in ((12, 31), (9, 30), (6, 30), (3, 31)):
            qe = date(yr, m, d)
            # 季末 + ~50 天後才大致公布完;未到就跳過
            if (today - qe).days >= 50:
                ends.append(qe.isoformat())
    ends = sorted(set(ends), reverse=True)
    return ends[:n]


def bulk(dataset, d):
    """date-bulk:不帶 data_id,抓某季『全台股』該 dataset。回傳 DataFrame(可能很大)。"""
    params = {"dataset": dataset, "date": d}
    if TOKEN:
        params["token"] = TOKEN
    r = requests.get(API, params=params, timeout=90)
    r.raise_for_status()
    js = r.json()
    return pd.DataFrame(js.get("data", []))


def pick_value(df, sid, type_names):
    """從 long 格式財報(欄:stock_id/type/value)取某股某科目值。"""
    sub = df[(df["stock_id"] == sid) & (df["type"].isin(type_names))]
    if sub.empty:
        return None
    return pd.to_numeric(sub["value"], errors="coerce").iloc[0]


def main():
    qends = recent_quarter_ends(EPS_HIGH_LOOKBACK + 1)
    if len(qends) < 4:
        print("可用季末不足 4 季,終止"); return
    print(f"掃描季末(最新在前):{qends[:4]} ...共 {len(qends)} 季")

    # 1) 全台股名稱/產業(1 call)
    try:
        info = bulk("TaiwanStockInfo", "")  # info 不需 date,但 API 容忍
    except Exception:
        info = pd.DataFrame()
    if info.empty or "stock_id" not in info.columns:
        # 退回:用 FinMind loader
        from FinMind.data import DataLoader
        dl = DataLoader()
        if TOKEN:
            dl.login_by_token(api_token=TOKEN)
        info = dl.taiwan_stock_info()
    namemap = {str(r["stock_id"]): str(r["stock_name"]) for _, r in info.iterrows()}
    indmap = {str(r["stock_id"]): str(r.get("industry_category", "")) for _, r in info.iterrows()}

    # 2) 近四季損益(EPS / 淨利)+ 歷史季EPS(算新高)— 逐季 bulk
    inc_by_q = {}
    for d in qends:
        try:
            inc_by_q[d] = bulk("TaiwanStockFinancialStatements", d)
        except Exception as e:
            print(f"  ⚠️ {d} 損益 bulk 失敗:{e}")
            inc_by_q[d] = pd.DataFrame()

    # 3) 最新季權益(算 ROE 分母)— 1 季 bulk
    bal = pd.DataFrame()
    for d in qends:
        try:
            b = bulk("TaiwanStockBalanceSheet", d)
            if not b.empty:
                bal = b
                break
        except Exception:
            continue

    # 以最新季有 EPS 的股票為母體
    last4 = qends[:4]
    base = inc_by_q.get(last4[0], pd.DataFrame())
    if base.empty:
        print("最新季損益抓不到,終止"); return
    sids = sorted(set(base["stock_id"].astype(str)))
    print(f"母體 {len(sids)} 檔,開始計算 ROE / EPS新高 ...")

    rows = []
    for sid in sids:
        # 近四季淨利加總
        ni4 = 0.0; have_ni = False
        for d in last4:
            v = pick_value(inc_by_q.get(d, pd.DataFrame()), sid,
                           ["IncomeAfterTaxes", "ProfitAfterTax", "NetIncome"])
            if v is not None and pd.notna(v):
                ni4 += float(v); have_ni = True
        eq = pick_value(bal, sid, ["Equity", "TotalEquity", "EquityAttributableToOwnersOfParent"])
        roe = (ni4 / eq * 100) if (have_ni and eq and eq > 0) else None

        # 近四季EPS 加總 vs 過去滾動四季最高(EPS新高)
        eps_q = []
        for d in qends:
            v = pick_value(inc_by_q.get(d, pd.DataFrame()), sid, ["EPS"])
            eps_q.append(float(v) if (v is not None and pd.notna(v)) else np.nan)
        ttm_now = np.nansum(eps_q[:4]) if np.isfinite(eps_q[:4]).any() else np.nan
        eps_high = False
        if pd.notna(ttm_now) and len(eps_q) >= 8:
            past_ttms = [np.nansum(eps_q[i:i+4]) for i in range(1, len(eps_q) - 3)]
            past_ttms = [x for x in past_ttms if pd.notna(x)]
            if past_ttms and ttm_now >= max(past_ttms):
                eps_high = True

        if (roe is not None and roe >= ROE_MIN) or eps_high:
            rows.append({
                "代號": sid, "名稱": namemap.get(sid, sid),
                "產業": indmap.get(sid, ""),
                "近四季ROE": round(roe, 1) if roe is not None else None,
                "近四季EPS": round(float(ttm_now), 2) if pd.notna(ttm_now) else None,
                "EPS新高": "✔" if eps_high else "",
            })

    df = pd.DataFrame(rows)
    if df.empty:
        print("無符合條件者(門檻可能太嚴或 bulk 無資料)"); return
    df = df.sort_values(["EPS新高", "近四季ROE"], ascending=[False, False], na_position="last")

    # 標出「不在現有 PICKS」的遺珠
    try:
        from fetch_fundamentals_tw import PICKS
        picks = set(PICKS)
    except Exception:
        picks = set()
    df["已在PICKS"] = df["代號"].apply(lambda s: "✔" if s in picks else "")
    new = df[df["已在PICKS"] == ""]

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(f"# 全市場獲利初篩  共 {len(df)} 檔(ROE≥{ROE_MIN} 或 EPS創近{EPS_HIGH_LOOKBACK}季高)\n")
        f.write(f"# 產生:{datetime.now().date().isoformat()}  其中不在現有PICKS(遺珠)={len(new)} 檔\n")
        f.write("# 代號  名稱  產業  ROE  EPS  EPS新高  已在PICKS\n")
        for _, r in df.iterrows():
            f.write(f"{r['代號']}  {r['名稱']}  {r['產業']}  "
                    f"ROE={r['近四季ROE']}  EPS={r['近四季EPS']}  "
                    f"{r['EPS新高']}  {r['已在PICKS']}\n")

    print(f"\n完成 → {OUT}")
    print(f"  符合 {len(df)} 檔,其中遺珠(不在PICKS){len(new)} 檔")
    print("\n遺珠 Top 20(ROE 高 / EPS新高,值得納入深度體檢):")
    show = new.head(20)[["代號", "名稱", "產業", "近四季ROE", "近四季EPS", "EPS新高"]]
    print(show.to_string(index=False))


if __name__ == "__main__":
    main()
