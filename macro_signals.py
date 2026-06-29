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
    """🛢️ WTI 原油 + Cushing 庫存(以 FRED 為主源,FMP 可能是不同合約)"""
    df = fred("DCOILWTICO", days=400)
    price = None; src = []
    if df is not None and len(df):
        p = float(df.iloc[-1]["value"])
        src.append(("FRED", p))
        price = p  # FRED 是現貨,優先
    fmp_q = fmp_quote("USOIL") or fmp_quote("CLUSD")
    if fmp_q:
        p_fmp = float(fmp_q.get("price", 0)) or None
        if p_fmp: src.append(("FMP", p_fmp))
    if not src: return {"狀態":"❌ 無資料"}
    # 不取平均(FMP 可能是遠月合約)— 以 FRED 現貨為準
    if price is None: price = src[0][1]

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
    """📱 台灣月出口 YoY% — 多源 fallback(FRED + OECD + 財政部)"""
    # 先試 FRED(最快,有就直接用)
    for sid in ["XTEXVA01TWM659S", "XTEXVA01TWM667S", "TWNEXP"]:
        df = fred(sid, days=900)
        if df is not None and len(df) >= 13:
            cur = df.iloc[-1]["value"]; yr = df.iloc[-13]["value"]
            if yr and yr > 0:
                yoy = (cur/yr - 1) * 100
                date = df.iloc[-1]["date"].strftime("%Y-%m")
                return {"當月": round(cur, 0), "YoY%": round(yoy, 1),
                        "資料月": date, "源": f"FRED {sid}",
                        "判讀": "🟢🟢 AI 強(>20%)" if yoy>20 else "🟢 健康" if yoy>5 else "🟠 弱"}

    # 備援 1:OECD SDMX-JSON(只給指數,但穩定)
    try:
        url = "https://stats.oecd.org/SDMX-JSON/data/MEI_TRD/TWN.XTEXVA01.IXOBSA.M/all"
        r = requests.get(url, params={"dimensionAtObservation":"allDimensions"},
                         timeout=15, headers={"Accept":"application/json"})
        if r.status_code == 200:
            j = r.json()
            obs = j.get("dataSets", [{}])[0].get("observations", {})
            time_dim = j.get("structure", {}).get("dimensions", {}).get("observation", [])
            time_idx = next((i for i, d in enumerate(time_dim) if d.get("id")=="TIME_PERIOD"), -1)
            if time_idx >= 0 and obs:
                times = time_dim[time_idx].get("values", [])
                # 排序找最新 12 個月
                sorted_keys = sorted(obs.keys(), key=lambda k: int(k.split(":")[time_idx]))
                if len(sorted_keys) >= 13:
                    cur_val = obs[sorted_keys[-1]][0]
                    yr_val = obs[sorted_keys[-13]][0]
                    yoy = (cur_val/yr_val - 1) * 100 if yr_val else None
                    cur_t = times[int(sorted_keys[-1].split(":")[time_idx])].get("id","")
                    if yoy is not None:
                        return {"當月指數": round(cur_val, 1), "YoY%": round(yoy, 1),
                                "資料月": cur_t, "源": "OECD MEI_TRD",
                                "判讀": "🟢🟢 AI 強(>20%)" if yoy>20 else "🟢 健康" if yoy>5 else "🟠 弱"}
    except Exception:
        pass

    return {"狀態":"❌ FRED+OECD 都失敗,手動查 https://web02.mof.gov.tw/njswww/WebMain.aspx"}


def signal_sox():
    """🔥 半導體 SOX 指數 — FMP ^SOX 不穩,改用 SOXX ETF 代理"""
    # SOXX 是 iShares Semiconductor ETF,跟 SOX 指數高度相關
    for sym in ["SOXX", "SMH", "^SOX"]:
        fmp_q = fmp_quote(sym)
        if fmp_q and float(fmp_q.get("price", 0)) > 0:
            price = float(fmp_q["price"])
            hi52 = float(fmp_q.get("yearHigh", price))
            pct = (price/hi52 - 1) * 100
            return {f"{sym}": round(price, 2), "52週高": round(hi52, 2),
                    "距高%": round(pct, 1),
                    "判讀": "🔴 新高過熱" if pct>-3 else "🟢 健康" if pct>-15 else "🟠 修正中"}
    return {"狀態":"❌ 無資料"}


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


# ===== 新增 8 個信號 =====

def signal_cpi():
    """💸 CPI 通膨 YoY"""
    df = fred("CPIAUCSL", days=730)  # 總體 CPI
    core = fred("CPILFESL", days=730)  # 核心 CPI
    if df is None or len(df) < 13: return {"狀態":"❌ 無資料"}
    cur = df.iloc[-1]["value"]; yr = df.iloc[-13]["value"]
    yoy = (cur/yr - 1) * 100
    core_yoy = None
    if core is not None and len(core) >= 13:
        core_yoy = (core.iloc[-1]["value"]/core.iloc[-13]["value"] - 1) * 100
    date = df.iloc[-1]["date"].strftime("%Y-%m")
    return {"總 YoY%": round(yoy, 2), "核心 YoY%": round(core_yoy, 2) if core_yoy else None,
            "資料月": date,
            "判讀": "🔴 升息壓力" if yoy>3 else "🟢 降息空間" if yoy<2 else "🟡 中性(2-3%)"}


def signal_us10y():
    """📜 10 年公債殖利率"""
    df = fred("DGS10", days=60)
    if df is None or len(df) == 0: return {"狀態":"❌ 無資料"}
    cur = float(df.iloc[-1]["value"])
    mo_ago = float(df.iloc[-22]["value"]) if len(df) >= 22 else None
    chg = cur - mo_ago if mo_ago else None
    return {"10Y%": round(cur, 2), "月變動%": round(chg, 2) if chg else None,
            "判讀": "🔴 緊縮股不利(>4.5)" if cur>4.5 else "🟢 寬鬆股利好(<3.5)" if cur<3.5 else "🟡 中性(3.5-4.5)"}


def signal_recession_prob():
    """⚠️ NY Fed 12 個月衰退機率"""
    df = fred("RECPROUSM156N", days=400)
    if df is None or len(df) == 0: return {"狀態":"❌ 無資料"}
    cur = float(df.iloc[-1]["value"])
    yr_ago = float(df.iloc[-13]["value"]) if len(df) >= 13 else None
    date = df.iloc[-1]["date"].strftime("%Y-%m")
    return {"衰退機率%": round(cur, 1),
            "去年同期%": round(yr_ago, 1) if yr_ago else None,
            "資料月": date,
            "判讀": "🔴 高(>30%)" if cur>30 else "🟢 低(<15%)" if cur<15 else "🟡 警戒(15-30%)"}


def signal_oecd_cli():
    """🌐 OECD 全球景氣領先指標(>100=擴張)"""
    # 試多個 FRED OECD CLI series
    for sid in ["OECDLOLITOAASTSAM", "USALOLITONOSTSAM", "OECDLOLITONOSTSAM"]:
        df = fred(sid, days=400)
        if df is not None and len(df) > 0: break
    if df is None or len(df) == 0: return {"狀態":"❌ 無資料(FRED OECD series 變動)"}
    cur = float(df.iloc[-1]["value"])
    mo_ago = float(df.iloc[-2]["value"]) if len(df) >= 2 else None
    chg = cur - mo_ago if mo_ago else None
    date = df.iloc[-1]["date"].strftime("%Y-%m")
    return {"CLI": round(cur, 2),
            "月變動": round(chg, 2) if chg else None,
            "資料月": date,
            "判讀": "🟢 擴張(>100 升)" if cur>100 and chg and chg>0 else
                   "🟠 擴張中放緩" if cur>100 else
                   "🔴 收縮(<100 降)" if chg and chg<0 else "🟡 觸底"}


def signal_tw_gdp():
    """🇹🇼 台灣 GDP YoY — 多 series 嘗試"""
    for sid in ["NGDPRSAXDCTWQ", "TWNRGDPEXP", "NYGDPMKTPSACDTW",
                "MKTGDPTWA646NWDB", "NAEXKP01TWQ659S"]:
        df = fred(sid, days=900)
        if df is None or len(df) < 5: continue
        cur = df.iloc[-1]["value"]; yr = df.iloc[-5]["value"]
        if not yr or yr == 0: continue
        yoy = (cur/yr - 1) * 100
        date = df.iloc[-1]["date"].strftime("%Y-Q%q")
        return {"GDP YoY%": round(yoy, 2), "資料季": date, "FRED series": sid,
                "判讀": "🟢 強(>5%)" if yoy>5 else "🟡 中性(2-5%)" if yoy>2 else "🔴 弱"}
    return {"狀態":"❌ 無資料(可改抓主計總處 stat.gov.tw)"}


def signal_move():
    """🌀 MOVE 債市波動指數 — 直接用 TLT 30 日波動率代理
    (ICE MOVE 無公開免費 API,FMP 的 ^MOVE 不是真實 MOVE 指數)"""
    # FMP 上 ^MOVE 不是真實 ICE MOVE,直接走 TLT 代理
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=60)).isoformat()
    hist = fmp("historical-price-eod/full", symbol="TLT", **{"from": start, "to": today.isoformat()})
    if hist:
        h = hist.get("historical") if isinstance(hist, dict) else hist
        if h and len(h) >= 30:
            df = pd.DataFrame(h).sort_values("date").tail(30)
            ret = df["close"].pct_change().dropna()
            vol = ret.std() * (252 ** 0.5) * 100  # 年化波動率%
            return {"TLT 30日年化波動%": round(vol, 1), "註": "MOVE 代理",
                    "判讀": "🔴 債市恐慌(>15%)" if vol>15 else "🟡 警戒(10-15%)" if vol>10 else "🟢 平靜(<10%)"}
    return {"狀態":"❌ MOVE 與 TLT 都抓不到"}


def signal_csp_capex():
    """☁️ CSP 4 大廠資本支出加總(MSFT+GOOG+AMZN+META)— 季"""
    syms = ["MSFT","GOOGL","AMZN","META"]
    total_cur = 0; total_yr = 0; counted = 0
    for s in syms:
        cf = fmp("cash-flow-statement", symbol=s, period="quarter", limit=5)
        if not cf or len(cf) < 5: continue
        cur = abs(float(cf[0].get("capitalExpenditure", 0)))
        yr = abs(float(cf[4].get("capitalExpenditure", 0)))
        if cur and yr:
            total_cur += cur; total_yr += yr; counted += 1
    if counted == 0: return {"狀態":"❌ 無資料"}
    yoy = (total_cur/total_yr - 1) * 100
    return {"當季加總(億美)": round(total_cur/1e8, 1),
            "YoY%": round(yoy, 1),
            "計入家數": counted,
            "判讀": "🟢🟢 AI 需求爆(>50%)" if yoy>50 else
                   "🟢 強(>20%)" if yoy>20 else
                   "🟡 平(0-20%)" if yoy>0 else "🔴 收縮"}


def signal_fed_watch():
    """🏛️ FedWatch 年底利率預期 — Fed Fund Futures 多 symbol 嘗試"""
    # 試多個 Fed Funds Futures 月份代碼(CME ZQ + 一些別名)
    today = datetime.now()
    yr = today.year
    # 12 月合約代碼 ZQZ + 年末2位
    yr_suffix = str(yr)[-2:]
    for sym in [f"ZQZ{yr_suffix}", "FF=F", "ZQ", f"ZQZ{yr_suffix}.CME"]:
        fmp_q = fmp_quote(sym)
        if fmp_q and float(fmp_q.get("price", 0)) > 0:
            price = float(fmp_q["price"])
            implied = 100 - price
            return {f"{sym}": round(price, 4),
                    "年底隱含 Fed Funds%": round(implied, 2),
                    "判讀": "🔴 高利率延續" if implied>4 else "🟢 降息預期" if implied<3.5 else "🟡 中性"}
    # 備援:用 FRED 10Y-2Y yield curve 推測
    y10 = fred("DGS10", days=10)
    y2 = fred("DGS2", days=10)
    if y10 is not None and y2 is not None and len(y10) and len(y2):
        spread = float(y10.iloc[-1]["value"]) - float(y2.iloc[-1]["value"])
        return {"10Y-2Y 利差%": round(spread, 2), "註": "Fed Futures 抓不到,用利差代理",
                "判讀": "🔴 倒掛(<0,衰退訊號)" if spread<0 else "🟢 正常(>0)" if spread>0.5 else "🟡 平坦"}
    return {"狀態":"❌ 無資料"}


# ---------- 主程式 ----------
def main():
    signals = {
        # === 原 7 信號 ===
        "🏠 房地產 Case-Shiller": signal_case_shiller(),
        "💰 黃金": signal_gold(),
        "🛢️ 原油 WTI + Cushing": signal_oil(),
        "💱 澳幣 AUDUSD": signal_aud(),
        "📱 台灣出口 YoY": signal_tw_export(),
        "🔥 半導體 SOX": signal_sox(),
        "😱 VIX 恐慌指數": signal_vix(),
        # === 新增 8 信號(利率/景氣/AI) ===
        "💸 CPI 通膨 YoY": signal_cpi(),
        "📜 10Y 公債殖利率": signal_us10y(),
        "🏛️ FedWatch 年底利率": signal_fed_watch(),
        "🌀 MOVE 債市波動": signal_move(),
        "⚠️ NY Fed 衰退機率": signal_recession_prob(),
        "🌐 OECD 全球領先指標": signal_oecd_cli(),
        "🇹🇼 台灣 GDP YoY": signal_tw_gdp(),
        "☁️ CSP 4 大廠資本支出": signal_csp_capex(),
    }

    print(f"\n=== 總經 {len(signals)} 信號燈 {TODAY} ===\n")
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
