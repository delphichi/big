# -*- coding: utf-8 -*-
"""
台股總經儀表板 tw_macro_dashboard.py
=======================================================================
一支腳本拉 6 個總經訊號, 整合成 macro_signals.xlsx 的台股版:

  1. 景氣對策信號 (TaiwanBusinessIndicator) - 月燈號
  2. 台幣 USD/EUR/JPY/CNY 匯率 (TaiwanExchangeRate)
  3. 央行利率 FED/BOJ/ECB/PBOC (InterestRate)
  4. 美債殖利率曲線 1M~30Y (GovernmentBondsYield)
  5. CNN 恐懼貪婪指數 (CnnFearGreedIndex)
  6. 台指 VIX (TaiwanOptionVix)

輸出 data/台股_總經儀表板.xlsx, 7 個分頁:
  - 總覽 (每信號最新值 + 判讀)
  - 景氣信號 (月度歷史)
  - 匯率 / 央行利率 / 美債曲線 / 恐懼貪婪 / 台指VIX (每個一頁日線)
"""
import os, time, requests
import pandas as pd
from datetime import datetime, timedelta

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股_總經儀表板.xlsx"
END = datetime.now().strftime("%Y-%m-%d")
START_3Y = (datetime.now() - timedelta(days=365*3)).strftime("%Y-%m-%d")
START_1Y = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
START_90D = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")


def fm(dataset, data_id=None, start=None, end=None):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for _ in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=30)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception:
            time.sleep(1)
    return pd.DataFrame()


def biz_signal_judge(score):
    """景氣信號分數 (9-16: 紅燈過熱, 17-22: 黃紅, 23-31: 綠燈穩, 32-37: 黃藍, 38-45: 藍燈衰退) - 分數 = monitoring"""
    s = score
    if s is None or pd.isna(s): return "—"
    s = int(s)
    if s >= 38: return "🔵藍燈(衰退)"
    if s >= 32: return "🟡黃藍燈"
    if s >= 23: return "🟢綠燈(穩定)"
    if s >= 17: return "🟠黃紅燈"
    return "🔴紅燈(過熱)"


def vix_judge(v):
    if v is None: return "—"
    if v > 25: return "🔴恐慌"
    if v > 20: return "🟡警戒"
    if v > 15: return "🟢平靜"
    return "🟢極平靜"


def fg_judge(v):
    if v is None: return "—"
    if v <= 25: return "🔵極度恐懼"
    if v <= 45: return "🟢恐懼(便宜)"
    if v <= 55: return "🟡中性"
    if v <= 75: return "🟠貪婪"
    return "🔴極度貪婪(逃)"


def fx_judge(c, v):
    """匯率判讀: USD 升值對台股壓力 / JPY 弱影響日廠 / CNY 升值利好 export"""
    if v is None: return "—"
    if c == "USD":
        if v > 32: return "🔴台幣超弱"
        if v > 31: return "🟡台幣弱"
        if v < 29: return "🟢台幣強"
        return "🟢正常"
    return ""


def yield_curve_judge(y2, y10):
    if y2 is None or y10 is None: return "—"
    spread = y10 - y2
    if spread < -0.5: return "🔴深度倒掛"
    if spread < 0: return "🟠倒掛(衰退訊號)"
    if spread < 0.5: return "🟡接近倒掛"
    return "🟢正常"


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN")
    print(f"台股總經儀表板 — 更新 {END}")
    rows = []
    sheets = {}

    # ─── 1. 景氣對策信號 (月) ───
    print("抓景氣對策信號...")
    biz = fm("TaiwanBusinessIndicator", start="2022-01-01")
    if not biz.empty:
        biz = biz.sort_values("date")
        last = biz.iloc[-1]
        score = last.get("monitoring")
        rows.append({
            "信號": "📊 景氣對策",
            "數值": score,
            "額外": f"色={last.get('monitoring_color')} 先行={last.get('leading')}",
            "判讀": biz_signal_judge(score),
            "更新日期": last.get("date"),
        })
        sheets["景氣信號"] = biz.tail(36)

    # ─── 2. 匯率 (USD/JPY/CNY/EUR 取 last 90d) ───
    fx_all = []
    for c in ["USD","JPY","CNY","EUR"]:
        print(f"抓 {c} 匯率...")
        fx = fm("TaiwanExchangeRate", data_id=c, start=START_90D)
        if not fx.empty:
            fx = fx.sort_values("date")
            last = fx.iloc[-1]
            sb = last.get("spot_buy")
            rows.append({
                "信號": f"💱 {c}/TWD",
                "數值": round(float(sb), 3) if sb else None,
                "額外": f"30d 前 = {round(float(fx.iloc[max(0,len(fx)-30)].get('spot_buy', sb)), 3)}",
                "判讀": fx_judge(c, float(sb)) if sb else "—",
                "更新日期": last.get("date"),
            })
            fx_all.append(fx.assign(__cur=c))
    if fx_all:
        sheets["匯率"] = pd.concat(fx_all, ignore_index=True)

    # ─── 3. 央行利率 ───
    ir_all = []
    for cb in ["FED","BOJ","ECB","PBOC"]:
        print(f"抓 {cb} 利率...")
        ir = fm("InterestRate", data_id=cb, start="2022-01-01")
        if not ir.empty:
            ir = ir.sort_values("date")
            last = ir.iloc[-1]
            v = last.get("interest_rate")
            rows.append({
                "信號": f"🏦 {cb} 利率",
                "數值": round(float(v), 3) if v is not None else None,
                "額外": last.get("full_country_name", ""),
                "判讀": "—",
                "更新日期": last.get("date"),
            })
            ir_all.append(ir.assign(__cb=cb))
    if ir_all:
        sheets["央行利率"] = pd.concat(ir_all, ignore_index=True)

    # ─── 4. 美債殖利率 (主要期別 + 10Y-2Y 利差) ───
    print("抓美債殖利率...")
    bond_all = []
    bond_latest = {}
    for tenor in ["United States 3-Month","United States 2-Year","United States 5-Year",
                  "United States 10-Year","United States 30-Year"]:
        b = fm("GovernmentBondsYield", data_id=tenor, start=START_90D)
        if not b.empty:
            b = b.sort_values("date")
            last = b.iloc[-1]
            v = last.get("value")
            short = tenor.replace("United States ", "")
            bond_latest[short] = round(float(v), 3) if v is not None else None
            rows.append({
                "信號": f"📉 美債 {short}",
                "數值": bond_latest[short],
                "額外": "",
                "判讀": "—",
                "更新日期": last.get("date"),
            })
            bond_all.append(b.assign(__tenor=short))
    # 10Y-2Y 利差
    if "10-Year" in bond_latest and "2-Year" in bond_latest:
        spread = bond_latest["10-Year"] - bond_latest["2-Year"]
        rows.append({
            "信號": "🚨 10Y-2Y 利差",
            "數值": round(spread, 3),
            "額外": f"10Y={bond_latest['10-Year']} 2Y={bond_latest['2-Year']}",
            "判讀": yield_curve_judge(bond_latest["2-Year"], bond_latest["10-Year"]),
            "更新日期": END,
        })
    if bond_all:
        sheets["美債曲線"] = pd.concat(bond_all, ignore_index=True)

    # ─── 5. CNN 恐懼貪婪 ───
    print("抓 CNN 恐懼貪婪...")
    fg = fm("CnnFearGreedIndex", start=START_90D)
    if not fg.empty:
        fg = fg.sort_values("date")
        last = fg.iloc[-1]
        v = last.get("fear_greed")
        rows.append({
            "信號": "😨 CNN 恐懼貪婪",
            "數值": round(float(v), 1) if v is not None else None,
            "額外": last.get("fear_greed_emotion", ""),
            "判讀": fg_judge(float(v)) if v is not None else "—",
            "更新日期": last.get("date"),
        })
        sheets["恐懼貪婪"] = fg.tail(90)

    # ─── 6. 台指 VIX ───
    print("抓台指 VIX...")
    vix = fm("TaiwanOptionVix", start=START_90D)
    if not vix.empty:
        vix = vix.sort_values(["date","time"] if "time" in vix.columns else ["date"])
        last = vix.iloc[-1]
        v = last.get("vix")
        rows.append({
            "信號": "📊 台指 VIX",
            "數值": round(float(v), 2) if v is not None else None,
            "額外": last.get("time", ""),
            "判讀": vix_judge(float(v)) if v is not None else "—",
            "更新日期": last.get("date"),
        })
        sheets["台指VIX"] = vix.tail(500)

    # ─── 7. 台灣資通訊出口 (全球科技股 3 週領先指標) ───
    # 用既有 tw_export_fetcher.py 三層 fallback (財政部 → OECD → 新聞)
    print("抓台灣資通訊出口...")
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("tw_export_fetcher", "tw_export_fetcher.py")
        m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
        exp_results = m.fetch_all() or []
        for d in exp_results[:1]:  # 只取主源
            num = d.get("出口總值(百萬美元)") or d.get("出口指數(2015=100)")
            rows.append({
                "信號": "🚢 台灣出口 ICT",
                "數值": num,
                "額外": f"{d.get('源','')} / {d.get('月份','')}",
                "判讀": "🟢領先指標(全球科技 3 週)",
                "更新日期": d.get("月份", "—"),
            })
    except Exception as e:
        print(f"  ⚠️ ICT 出口失敗: {e}")

    # ─── 輸出 ───
    df = pd.DataFrame(rows)
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="總覽", index=False)
        for name, d in sheets.items():
            d.to_excel(xw, sheet_name=name[:31], index=False)

    print(f"\n→ {DST}")
    print(f"\n=== 總覽 ===")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
