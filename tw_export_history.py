# -*- coding: utf-8 -*-
"""
台灣資通訊出口歷史比對 tw_export_history.py
=======================================================================
拉 10 年台灣出口指數歷史, 現在 vs 5y 均 / 5y 高 / 10y 高
資料源多重 fallback:
  1. FRED XTEXVA01TWM659S (Taiwan Exports SA)
  2. OECD SDMX (Taiwan Exports Value)
  3. FinMind (無專屬 dataset, 略過)

跑法:
  FRED_API_KEY=xxx python tw_export_history.py

輸出:
  console 顯示 + data/tw_export_history.xlsx
"""
import os, requests, sys
import pandas as pd
from datetime import datetime, timedelta

FRED_KEY = os.environ.get("FRED_API_KEY", "")
DST = "data/tw_export_history.xlsx"


def fred(series_id, days=3650):
    if not FRED_KEY: return None
    url = "https://api.stlouisfed.org/fred/series/observations"
    start = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    try:
        r = requests.get(url, params={"series_id":series_id, "api_key":FRED_KEY,
                                       "file_type":"json", "observation_start":start},
                         timeout=20)
        if r.status_code != 200: return None
        obs = r.json().get("observations", [])
        rows = [(o["date"], float(o["value"])) for o in obs if o.get("value") not in (".","")]
        return pd.DataFrame(rows, columns=["date","value"])
    except Exception: return None


def oecd():
    url = "https://stats.oecd.org/SDMX-JSON/data/MEI_TRD/TWN.XTEXVA01.IXOBSA.M/all"
    try:
        r = requests.get(url, params={"dimensionAtObservation":"allDimensions"},
                         timeout=30, headers={"Accept":"application/json"})
        if r.status_code != 200: return None
        j = r.json()
        obs = j.get("dataSets",[{}])[0].get("observations",{})
        dims = j.get("structure",{}).get("dimensions",{}).get("observation",[])
        time_idx = next((i for i,d in enumerate(dims) if d.get("id")=="TIME_PERIOD"), -1)
        times = dims[time_idx].get("values",[]) if time_idx>=0 else []
        rows = []
        for key, val in obs.items():
            t_i = int(key.split(":")[time_idx])
            t = times[t_i].get("id","")
            v = val[0] if val else None
            if v is not None: rows.append((t, float(v)))
        rows.sort()
        return pd.DataFrame(rows, columns=["date","value"])
    except Exception: return None


def analyze(df, label, unit=""):
    if df is None or len(df) < 60: return None
    df = df.sort_values("date").reset_index(drop=True)
    latest = df.iloc[-1]["value"]
    latest_date = df.iloc[-1]["date"]

    # 5y / 10y 區間
    df["dt"] = pd.to_datetime(df["date"])
    now = df["dt"].max()
    v5 = df[df["dt"] >= now - pd.DateOffset(years=5)]["value"]
    v10 = df[df["dt"] >= now - pd.DateOffset(years=10)]["value"]

    # YoY (12 個月前對比)
    yoy = None
    if len(df) >= 13:
        prev = df.iloc[-13]["value"]
        if prev > 0: yoy = round((latest/prev - 1)*100, 1)

    print(f"\n{'='*70}\n{label}\n{'='*70}")
    print(f"最新: {latest:.1f}{unit}  ({latest_date})")
    print(f"YoY: {yoy:+.1f}%" if yoy is not None else "YoY: —")
    print(f"5y 均: {v5.mean():.1f}{unit}  高: {v5.max():.1f}  低: {v5.min():.1f}")
    print(f"10y 均: {v10.mean():.1f}{unit}  高: {v10.max():.1f}  低: {v10.min():.1f}")
    print(f"vs 5y 均: {(latest/v5.mean()-1)*100:+.1f}%")
    print(f"vs 5y 高: {(latest/v5.max()-1)*100:+.1f}%")
    print(f"vs 10y 高: {(latest/v10.max()-1)*100:+.1f}%")

    # 分位判讀
    pctile = (v5 < latest).mean() * 100
    print(f"位於 5y 分布 {pctile:.0f} 百分位")
    if pctile > 95: verdict = "🔴 歷史極高(過熱)"
    elif pctile > 80: verdict = "🟠 偏高"
    elif pctile > 50: verdict = "🟡 中偏高"
    elif pctile > 20: verdict = "🟢 中性"
    else: verdict = "🔵 偏低"
    print(f"判讀: {verdict}")

    return {"label":label, "最新":latest, "日期":latest_date, "YoY%":yoy,
            "5y均":round(v5.mean(),1), "5y高":round(v5.max(),1),
            "vs5y均%":round((latest/v5.mean()-1)*100,1),
            "vs5y高%":round((latest/v5.max()-1)*100,1),
            "5y分位":round(pctile,0), "判讀":verdict, "df":df}


def main():
    print(f"=== 台灣出口歷史比對 ({datetime.now().strftime('%Y-%m-%d')}) ===")
    results = []

    # 試 FRED 多個 series
    for sid, label in [
        ("XTEXVA01TWM659S", "FRED Taiwan Exports Value SA (百萬 USD)"),
        ("XTEXVA01TWM667S", "FRED Taiwan Exports Value SA Alt"),
        ("XTEITT01TWM667N", "FRED Taiwan Exports of ICT"),
    ]:
        df = fred(sid)
        if df is not None and len(df) > 24:
            r = analyze(df, label, unit=" 百萬")
            if r: results.append(r); break

    # 試 OECD (指數 2015=100)
    df = oecd()
    if df is not None:
        r = analyze(df, "OECD Taiwan Exports Index (2015=100)")
        if r: results.append(r)

    if not results:
        print("\n⚠️ 三源都無法連(可能網路限制或 FRED_API_KEY 未設)")
        print("解: 去 GitHub Actions 觸發此 workflow, 或本機執行")
        return

    # 存 xlsx
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        summary = pd.DataFrame([{k:v for k,v in r.items() if k != "df"} for r in results])
        summary.to_excel(xw, sheet_name="總覽", index=False)
        for r in results:
            df = r["df"][["date","value"]].tail(120)
            df.to_excel(xw, sheet_name=r["label"][:30], index=False)
    print(f"\n→ {DST}")


if __name__ == "__main__":
    main()
