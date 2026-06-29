# -*- coding: utf-8 -*-
"""
總經 7 信號燈 macro_signals.py
=======================================================================
每週日跑,抓 7 個信號燈 + 三源交叉驗證 + 寄 email
信號:
  🏠 房地產(Case-Shiller YoY)
  💰 黃金($/oz)
  🛢️ 原油 WTI($/bbl + Cushing 庫存)
  💱 澳幣(AUDUSD)
  📱 台灣月出口 YoY%
  🔥 半導體(SOX 指數 vs 52 週高)
  😱 VIX(恐慌指數)

API 需求:
  FMP_API_KEY  - 已有
  FRED_API_KEY - 免費註冊 fred.stlouisfed.org
  EIA_API_KEY  - 免費註冊 eia.gov/opendata
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

FMP_KEY = os.environ.get("FMP_API_KEY", "")
FRED_KEY = os.environ.get("FRED_API_KEY", "")
EIA_KEY = os.environ.get("EIA_API_KEY", "")

FMP_BASE = "https://financialmodelingprep.com/stable"
FRED_BASE = "https://api.stlouisfed.org/fred/series/observations"
EIA_BASE = "https://api.eia.gov/v2"

TODAY = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d")
DST = "data/macro_signals.xlsx"
LOG = "data/macro_signals_log.csv"


# ---------- 資料抓取 ----------
def fred(series_id, days=400):
    """FRED 通用抓取"""
    if not FRED_KEY: return None
    end = datetime.now().date()
    start = (end - timedelta(days=days)).isoformat()
    try:
        r = requests.get(FRED_BASE, params={
            "series_id": series_id, "api_key": FRED_KEY, "file_type": "json",
            "observation_start": start
        }, timeout=15)
        if r.status_code != 200: return None
        obs = r.json().get("observations", [])
        df = pd.DataFrame(obs)
        df = df[df["value"] != "."].copy()
        df["date"] = pd.to_datetime(df["date"])
        df["value"] = df["value"].astype(float)
        return df.sort_values("date").reset_index(drop=True)
    except Exception:
        return None


def fmp(endpoint, **params):
    if not FMP_KEY: return None
    params["apikey"] = FMP_KEY
    try:
        r = requests.get(f"{FMP_BASE}/{endpoint}", params=params, timeout=15)
        if r.status_code != 200: return None
        return r.json()
    except Exception:
        return None


def fmp_quote(sym):
    d = fmp("quote", symbol=sym)
    if isinstance(d, list) and d: return d[0]
    return None


# ---------- 7 個信號 ----------
def signal_case_shiller():
    """🏠 房地產 Case-Shiller 全美房價指數 YoY%"""
    df = fred("CSUSHPISA", days=730)
    if df is None or len(df) < 13: return {"狀態":"❌ 無資料"}
    cur = df.iloc[-1]["value"]; yr = df.iloc[-13]["value"]
    yoy = (cur/yr - 1) * 100
    date = df.iloc[-1]["date"].strftime("%Y-%m")
    return {"當前": round(cur,1), "去年同期": round(yr,1),
            "YoY%": round(yoy,2), "資料月": date,
            "判讀": "🔴 過熱" if yoy>10 else "🟢 平穩" if yoy>0 else "🟠 修正中"}


def signal_gold():
    """💰 黃金 $/oz"""
    fred_df = fred("GOLDAMGBD228NLBM", days=400)
    fmp_q = fmp_quote("GCUSD") or fmp_quote("GC")
    price = None; src = []
    if fred_df is not None and len(fred_df):
        p_fred = float(fred_df.iloc[-1]["value"])
        src.append(("FRED", p_fred))
    if fmp_q:
        p_fmp = float(fmp_q.get("price", 0)) or None
        if p_fmp: src.append(("FMP", p_fmp))
    if not src: return {"狀態":"❌ 無資料"}
    prices = [p for _, p in src]
    price = sum(prices) / len(prices)
    diff = max(prices) - min(prices) if len(prices) > 1 else 0
    return {"價格": round(price, 1), "資料源": ", ".join(f"{k}={v:.1f}" for k,v in src),
            "誤差": round(diff, 1),
            "判讀": "🟢 強(央行買盤)" if price>2400 else "🟡 區間" if price>2000 else "🔴 弱"}


def signal_oil():
    """🛢️ WTI 原油 + Cushing 庫存"""
    # FRED WTI 現貨
    df = fred("DCOILWTICO", days=400)
    price = None; src = []
    if df is not None and len(df):
        p = float(df.iloc[-1]["value"])
        src.append(("FRED", p))
    fmp_q = fmp_quote("CLUSD") or fmp_quote("CL")
    if fmp_q:
        p_fmp = float(fmp_q.get("price", 0)) or None
        if p_fmp: src.append(("FMP", p_fmp))
    if not src: return {"狀態":"❌ 無資料"}
    prices = [p for _, p in src]
    price = sum(prices) / len(prices)

    # Cushing 庫存(EIA)— 週度
    cushing = None
    if EIA_KEY:
        try:
            r = requests.get(f"{EIA_BASE}/petroleum/stoc/wstk/data/", params={
                "api_key": EIA_KEY, "frequency": "weekly",
                "data[0]": "value", "facets[series][]": "WCESTUS1",
                "sort[0][column]": "period", "sort[0][direction]": "desc",
                "length": 1
            }, timeout=15)
            j = r.json().get("response", {}).get("data", [])
            if j: cushing = j[0].get("value")
        except Exception: pass

    return {"WTI": round(price, 2), "資料源": ", ".join(f"{k}={v:.2f}" for k,v in src),
            "Cushing 庫存(千桶)": cushing,
            "判讀": "🔴 過熱(>$90)" if price>90 else "🟢 區間($60-85)" if price>60 else "🟠 弱"}


def signal_aud():
    """💱 澳幣 AUDUSD"""
    df = fred("DEXUSAL", days=400)  # 注意 DEXUSAL = USD per AUD,即 AUDUSD
    src = []
    if df is not None and len(df):
        p = float(df.iloc[-1]["value"])
        src.append(("FRED", p))
    fmp_q = fmp_quote("AUDUSD")
    if fmp_q:
        p_fmp = float(fmp_q.get("price", 0)) or None
        if p_fmp: src.append(("FMP", p_fmp))
    if not src: return {"狀態":"❌ 無資料"}
    prices = [p for _, p in src]
    price = sum(prices) / len(prices)
    return {"AUDUSD": round(price, 4), "資料源": ", ".join(f"{k}={v:.4f}" for k,v in src),
            "判讀": "🟢 強(中國需求好)" if price>0.68 else "🟡 區間" if price>0.62 else "🔴 弱(中國降溫)"}


def signal_tw_export():
    """📱 台灣月出口 YoY% — 用 FRED proxy (USD value) 因財政部 API 較難整合"""
    # FRED 台灣出口 (USD millions, 月)
    df = fred("XTEXVA01TWM664S", days=900)  # 台灣出口商品總值
    if df is None or len(df) < 13:
        # 備援:FRED 簡化版
        df = fred("TWNXR", days=900)
    if df is None or len(df) < 13: return {"狀態":"❌ 無資料(可手動查財政部)"}
    cur = df.iloc[-1]["value"]
    yr = df.iloc[-13]["value"] if len(df) >= 13 else None
    yoy = (cur/yr - 1) * 100 if yr else None
    date = df.iloc[-1]["date"].strftime("%Y-%m")
    return {"當月": round(cur, 0), "YoY%": round(yoy, 1) if yoy else None,
            "資料月": date,
            "判讀": "🟢🟢 AI 強(>20%)" if yoy and yoy>20 else "🟢 健康" if yoy and yoy>5 else "🟠 弱"}


def signal_sox():
    """🔥 半導體 SOX 指數"""
    fmp_q = fmp_quote("^SOX")
    if not fmp_q: return {"狀態":"❌ 無資料"}
    price = float(fmp_q.get("price", 0))
    hi52 = float(fmp_q.get("yearHigh", price))
    lo52 = float(fmp_q.get("yearLow", price))
    pct = (price/hi52 - 1) * 100
    return {"SOX": round(price, 2), "52週高": round(hi52,2),
            "距高%": round(pct, 1),
            "判讀": "🔴 新高過熱" if pct>-3 else "🟢 健康" if pct>-15 else "🟠 修正中"}


def signal_vix():
    """😱 VIX 恐慌指數"""
    df = fred("VIXCLS", days=30)
    src = []
    if df is not None and len(df):
        p = float(df.iloc[-1]["value"])
        src.append(("FRED", p))
    fmp_q = fmp_quote("^VIX")
    if fmp_q:
        p_fmp = float(fmp_q.get("price", 0)) or None
        if p_fmp: src.append(("FMP", p_fmp))
    if not src: return {"狀態":"❌ 無資料"}
    prices = [p for _, p in src]
    price = sum(prices) / len(prices)
    return {"VIX": round(price, 2), "資料源": ", ".join(f"{k}={v:.2f}" for k,v in src),
            "判讀": "🔴 恐慌(>25)" if price>25 else "🟡 警戒(20-25)" if price>20 else "🟢 平靜(<20)"}


# ---------- 主程式 ----------
def main():
    signals = {
        "🏠 房地產 Case-Shiller": signal_case_shiller(),
        "💰 黃金": signal_gold(),
        "🛢️ 原油 WTI + Cushing": signal_oil(),
        "💱 澳幣 AUDUSD": signal_aud(),
        "📱 台灣出口 YoY": signal_tw_export(),
        "🔥 半導體 SOX": signal_sox(),
        "😱 VIX 恐慌指數": signal_vix(),
    }

    print(f"\n=== 總經 7 信號燈 {TODAY} ===\n")
    rows = []
    for name, data in signals.items():
        print(f"\n{name}")
        for k, v in data.items():
            print(f"  {k}: {v}")
        rows.append({"信號": name, **data})

    # 輸出 Excel + 歷史 CSV
    os.makedirs("data", exist_ok=True)
    pd.DataFrame(rows).to_excel(DST, index=False)
    log_row = {"日期": TODAY}
    for name, data in signals.items():
        for k, v in data.items():
            if k in ("判讀","狀態","資料源","誤差"): continue
            log_row[f"{name.split(' ')[0]}_{k}"] = v
    log_df = pd.DataFrame([log_row])
    if os.path.exists(LOG):
        old = pd.read_csv(LOG)
        log_df = pd.concat([old, log_df], ignore_index=True)
    log_df.to_csv(LOG, index=False)
    print(f"\n→ 已輸出 {DST}")
    print(f"→ 歷史紀錄 {LOG}")


if __name__ == "__main__":
    main()
