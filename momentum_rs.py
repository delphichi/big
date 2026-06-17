#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股資金面三指標  momentum_rs.py
================================
只用每日股價(不需財報),計算「主力資金是否正在持續買進」的三個角度:

  RS(相對強弱)   = 個股 ÷ 大盤 的比值線。上升=跑贏大盤;報「近半年相對報酬」與「RS是否創52週高」
  週斜率          = 最近13週對數收盤價的線性回歸斜率(每週複合成長率);另算「加速度」=近13週 vs 前13週
  共振分數        = 10個條件同時檢查的得分(0-100);全亮代表趨勢/動能/資金同方向

大盤基準預設用 0050(穩定);要用加權指數見 fetch_benchmark 註解。
資料源:FinMind taiwan_stock_daily。建議抓 2 年以上(算52週高與斜率)。

pip install finmind pandas numpy openpyxl
"""

import os, time
import numpy as np
import pandas as pd

TICKERS = ["2330", "2454", "2379", "3231", "2408"]   # 台積電 聯發科 瑞昱 緯創 南亞科
BENCHMARK = "0050"            # 大盤代理;可改 "TAIEX"(見 fetch_benchmark)
START_DATE = "2023-06-01"     # 取 2+ 年供 52 週高 / 斜率
TOKEN = os.environ.get("FINMIND_TOKEN", "")
OUTPUT = "台股資金面三指標.xlsx"

D_YEAR, D_13W, D_26W = 252, 65, 130   # 交易日:52週 / 13週 / 26週


# ---------- 取數 ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        dl.login_by_token(api_token=TOKEN)
    return dl

def fetch_price(dl, sid, start):
    df = dl.taiwan_stock_daily(stock_id=sid, start_date=start)
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").set_index("date")
    vol = "Trading_Volume" if "Trading_Volume" in df.columns else "volume"
    out = pd.DataFrame({"close": df["close"], "vol": df[vol]})
    return out[out["close"] > 0]

def fetch_benchmark(dl, start):
    # 預設:0050 當大盤代理(欄位標準、最穩)
    # 想用真正加權指數,改成:
    #   df = dl.taiwan_stock_total_return_index(index_id="TAIEX", start_date=start)
    #   再取其價格欄位即可
    b = fetch_price(dl, BENCHMARK, start)
    return b["close"] if not b.empty else pd.Series(dtype=float)


# ---------- 數學工具 ----------
def log_slope(prices):
    """對數價格線性回歸斜率 = 每期複合成長率。"""
    p = pd.to_numeric(prices, errors="coerce").dropna()
    if len(p) < 5:
        return None
    x = np.arange(len(p))
    return float(np.polyfit(x, np.log(p.values), 1)[0])


# ---------- 三指標 ----------
def compute(stock, bench):
    s = stock["close"]; v = stock["vol"]
    if len(s) < D_26W:
        return None
    out = {}

    # ── RS:個股/大盤 比值線 ──
    aligned = pd.concat([s.rename("s"), bench.rename("b")], axis=1).dropna()
    rs = (aligned["s"] / aligned["b"])
    out["近半年相對報酬%"] = None
    if len(aligned) > D_26W:
        s_ret = aligned["s"].iloc[-1] / aligned["s"].iloc[-D_26W] - 1
        b_ret = aligned["b"].iloc[-1] / aligned["b"].iloc[-D_26W] - 1
        out["近半年相對報酬%"] = round((s_ret - b_ret) * 100, 1)
    rs_newhigh = (len(rs) >= D_YEAR and rs.iloc[-1] >= rs.tail(D_YEAR).max() * 0.98)
    rs_rising  = (len(rs) > D_13W and rs.iloc[-1] > rs.iloc[-D_13W])
    out["RS創52週高"] = "✔" if rs_newhigh else ""
    out["RS上升"]    = "✔" if rs_rising else ""

    # ── 週斜率 + 加速度 ──
    wk = s.resample("W-FRI").last().dropna()
    slope_recent = log_slope(wk.tail(13))
    slope_prior  = log_slope(wk.iloc[-26:-13]) if len(wk) >= 26 else None
    out["週斜率%/週"] = round(slope_recent * 100, 2) if slope_recent is not None else None
    accel = (slope_recent is not None and slope_prior is not None and slope_recent > slope_prior)
    out["週斜率加速"] = "✔" if accel else ""
    wk_pos = (slope_recent is not None and slope_recent > 0)

    # ── 日斜率 ──
    d_slope = log_slope(s.tail(20))
    day_pos = (d_slope is not None and d_slope > 0)

    # ── 均線 / 創高 / 量能 ──
    ma50  = s.tail(50).mean();  ma200 = s.tail(200).mean()
    price = s.iloc[-1]
    above200 = price > ma200
    above50  = price > ma50
    ma_align = ma50 > ma200
    p_newhigh = (len(s) >= D_YEAR and price >= s.tail(D_YEAR).max() * 0.95)
    vol_surge = v.tail(20).mean() > v.tail(60).mean()

    # ── 大盤多頭 ──
    bench_up = (len(bench) >= 200 and bench.iloc[-1] > bench.tail(200).mean())

    # ── 共振分數(10 條) ──
    conds = {
        "站上200日線": above200, "站上50日線": above50, "均線多頭排列": ma_align,
        "價格創52週高": p_newhigh, "RS創52週高": rs_newhigh, "RS上升": rs_rising,
        "週斜率為正": wk_pos, "週斜率加速": accel, "量能放大": vol_surge, "大盤多頭": bench_up,
    }
    score = round(sum(bool(x) for x in conds.values()) / len(conds) * 100)
    out["共振分數"] = score
    out["亮燈"] = "／".join(k for k, x in conds.items() if x)
    out["收盤"] = round(price, 1)
    return out


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    bench = fetch_benchmark(dl, START_DATE)
    if bench.empty:
        print("⚠ 無法取得大盤基準,RS/大盤多頭將失準"); 
    rows = []
    for sid in TICKERS:
        print(f"分析 {sid} ...")
        try:
            stock = fetch_price(dl, sid, START_DATE)
            time.sleep(0.8)
            res = compute(stock, bench)
            if res:
                rows.append({"代號": sid, **res})
            else:
                print(f"  {sid} 資料不足")
        except Exception as e:
            print(f"  ! {sid} 失敗:{e}")

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("共振分數", ascending=False)
        cols = ["代號", "共振分數", "近半年相對報酬%", "RS創52週高", "RS上升",
                "週斜率%/週", "週斜率加速", "收盤", "亮燈"]
        df = df[[c for c in cols if c in df.columns]]
    df.to_excel(OUTPUT, sheet_name="資金面三指標", index=False)
    print(f"\n已輸出:{OUTPUT}\n")
    pd.set_option("display.unicode.east_asian_width", True); pd.set_option("display.width", 240)
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
