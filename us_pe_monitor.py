# -*- coding: utf-8 -*-
"""
美股 PE / PEG / Forward PE 監看表 us_pe_monitor.py
=======================================================================
監看清單(96 檔 + 自訂),抓即時 PE / PEG / Forward PE / 距合理價%
輸出 data/美股PE監看表.xlsx

跑法:
  python us_pe_monitor.py            # 用內建清單
  WATCHLIST=AAPL,NVDA python us_pe_monitor.py  # 自訂

每次跑會:
  1) 抓即時報價 + 分析師 EPS 預估
  2) 算當前 PE / Forward PE / PEG(用 EPS3y% 為 g)
  3) 跟前次快照比對(漲跌幅 / PE 變化)
  4) 標「便宜/合理/偏貴/過熱」鬧鐘
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
SRC = "data/美股體檢總表.xlsx"
DST = "data/美股PE監看表.xlsx"
WORKERS = int(os.environ.get("MONITOR_WORKERS", "6"))

DEFAULT_WATCH = """
LIN USLM WPM ECL META NFLX GOOG RACE DECK MELI PDD HIG ACGL PYPL KNSL PRDO
AXP BABA LLY NVO RMD DXCM NBIX EXEL UHS ADP WTKWY RELX HWM CYATY NEE CAT GLW
CLS BRK-B VRT PLTR NVDA MSFT IBM TSM AVGO ASML KLAC NVMI APH ANET AMZN ADBE
INTU PAYC PCTY QLYS IDCC NOW ADSK FTNT NTES CRM TMUS CCEP KOF COST WMT LITE
COHR MRVL AMD CIEN WES UAN FI BR NICE TRAK DSGX ACN INFY AAPL LRCX AMAT
CDNS GSK AMG NFG CF AER MU FSLR WDC SNDK TSLA MA V GLD
""".split()


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=15)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1)); continue
        if r.status_code != 200:
            return None
        try: return r.json()
        except: return None
    return None


def fetch_quote(sym):
    """抓即時報價 + 分析師預估明年 EPS"""
    try:
        q = get("quote", symbol=sym)
        if not q: return sym, None, None, None
        q0 = q[0] if isinstance(q, list) else q
        price = q0.get("price")
        pe_curr = q0.get("pe")  # FMP 提供當前 PE(TTM)

        est = get("analyst-estimates", symbol=sym, period="annual", limit=2)
        eps_next = None
        if isinstance(est, list) and est:
            # 取最近未來一年的 EPS 估值
            for e in est:
                v = e.get("estimatedEpsAvg") or e.get("epsAvg")
                if v and v > 0:
                    eps_next = v
                    break
        fwd_pe = round(price / eps_next, 1) if (price and eps_next and eps_next > 0) else None
        return sym, price, pe_curr, fwd_pe
    except Exception:
        return sym, None, None, None


def alarm(pe, fwd_pe, peg):
    """估值鬧鐘"""
    if peg is not None:
        if peg < 1.0: return "🟢未來便宜"
        if peg < 1.5: return "🟢成長未反映"
        if peg < 2.0: return "🟡未來合理"
        if peg < 3.0: return "🟠未來偏貴"
        return "🔴未來過熱"
    if fwd_pe is not None:
        if fwd_pe < 12: return "🟢未來便宜"
        if fwd_pe < 18: return "🟡未來合理"
        if fwd_pe < 30: return "🟠未來偏貴"
        return "🔴未來過熱"
    if pe is not None:
        if pe < 12: return "🟢便宜"
        if pe < 20: return "🟡合理"
        if pe < 35: return "🟠偏貴"
        return "🔴過熱"
    return "—"


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY"); return

    watch = os.environ.get("WATCHLIST", "").strip()
    syms = [s.strip().upper() for s in watch.split(",") if s.strip()] if watch else DEFAULT_WATCH
    print(f"監看 {len(syms)} 檔")

    # 從體檢總表撈基礎欄(品質/EPS3y/含金量/評等等)
    base = pd.read_excel(SRC, sheet_name="體檢總表")
    base["代號"] = base["代號"].astype(str)
    base_keep = ["代號","名稱","產業","評等","品質總分","EPS3y%","ROE%","含金量",
                 "營收CAGR%","循環股","主要漏洞","市值(億美)"]
    base_keep = [c for c in base_keep if c in base.columns]
    base = base[base["代號"].isin(syms)][base_keep]

    # 平行抓即時報價
    quotes = {}
    print(f"抓即時報價 ({WORKERS} 平行)...")
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_quote, s): s for s in syms}
        for fut in as_completed(futs):
            sym, price, pe, fwd_pe = fut.result()
            quotes[sym] = {"當前股價": price, "PER即時": pe, "ForwardPE即時": fwd_pe}

    # 組裝
    rows = []
    for s in syms:
        b = base[base["代號"] == s]
        b = b.iloc[0].to_dict() if len(b) else {"代號": s, "名稱": "(不在總表)"}
        q = quotes.get(s, {})
        b.update(q)
        # PEG = ForwardPE / EPS3y
        g = pd.to_numeric(b.get("EPS3y%"), errors="coerce")
        fwd = b.get("ForwardPE即時")
        peg = round(fwd / g, 2) if (fwd and g and g > 0) else None
        b["PEG即時"] = peg
        b["估值鬧鐘"] = alarm(b.get("PER即時"), fwd, peg)
        rows.append(b)

    df = pd.DataFrame(rows)
    # 排序:過熱 → 偏貴 → 合理 → 便宜
    order = {"🔴未來過熱":0, "🔴過熱":1, "🟠未來偏貴":2, "🟠偏貴":3,
             "🟡未來合理":4, "🟡合理":5, "🟢成長未反映":6, "🟢未來便宜":7, "🟢便宜":8, "—":9}
    df["_o"] = df["估值鬧鐘"].map(lambda x: order.get(x, 9))
    df = df.sort_values(["_o","品質總分"], ascending=[True, False]).drop(columns=["_o"])

    # 與前次快照比對(若存在)
    prev_path = DST
    if os.path.exists(prev_path):
        try:
            prev = pd.read_excel(prev_path, sheet_name="監看表")
            prev = prev[["代號","當前股價"]].rename(columns={"當前股價":"前次股價"})
            df = df.merge(prev, on="代號", how="left")
            df["漲跌%"] = ((df["當前股價"] - df["前次股價"]) / df["前次股價"] * 100).round(2)
        except Exception:
            pass

    # 欄位順序
    front = ["代號","名稱","產業","評等","品質總分","當前股價","PER即時",
             "ForwardPE即時","EPS3y%","PEG即時","估值鬧鐘"]
    if "漲跌%" in df.columns: front.append("漲跌%")
    if "前次股價" in df.columns: front.append("前次股價")
    rest = [c for c in df.columns if c not in front]
    df = df[front + rest]

    # 寫出
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="監看表", index=False)
        # 按估值鬧鐘分頁
        for alarm_type in ["🟢成長未反映","🟢未來便宜","🟢便宜"]:
            sub = df[df["估值鬧鐘"] == alarm_type]
            if len(sub):
                sheet = alarm_type.replace("🟢","").replace("🟡","").replace("🔴","").replace("🟠","")
                sub.to_excel(xw, sheet_name=f"買進_{sheet}"[:31], index=False)
        for alarm_type in ["🔴未來過熱","🔴過熱","🟠未來偏貴"]:
            sub = df[df["估值鬧鐘"] == alarm_type]
            if len(sub):
                sheet = alarm_type.replace("🟢","").replace("🟡","").replace("🔴","").replace("🟠","")
                sub.to_excel(xw, sheet_name=f"警示_{sheet}"[:31], index=False)

    # 摘要
    print(f"\n→ 已輸出 {DST}")
    print(f"\n估值鬧鐘分布:")
    print(df["估值鬧鐘"].value_counts().to_string())
    print(f"\n🟢 買進信號清單(成長未反映 / 未來便宜):")
    buy = df[df["估值鬧鐘"].isin(["🟢成長未反映","🟢未來便宜","🟢便宜"])]
    cols = ["代號","名稱","評等","當前股價","ForwardPE即時","PEG即時","估值鬧鐘"]
    print(buy[cols].head(20).to_string(index=False))


if __name__ == "__main__":
    main()
