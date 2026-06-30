# -*- coding: utf-8 -*-
"""
美股全景儀表板 us_full_dashboard.py
=======================================================================
一次性對 watchlist 跑 5 個 FMP 端點, 整合成一張多分頁總表:

  1. DCF 估值雷達      → discounted-cash-flow + levered-discounted-cash-flow
  2. 體質健診           → financial-scores (Altman Z + Piotroski)
  3. 產品/地理曝險      → revenue-product-segmentation + revenue-geographic-segmentation
  4. 內部人 + 國會交易   → insider-trading/statistics + senate-trades + house-trades
  5. Owner Earnings    → owner-earnings (Buffett 真實盈餘)

Watchlist 來源: TICKERS env → data/watchlist_us.txt → fallback

輸出 data/美股_全景儀表板.xlsx, 7 個分頁:
  - 總覽 (每檔一行, 五大模組分數 + 綜合訊號)
  - DCF 估值
  - 體質健診
  - 產品結構
  - 地理結構
  - 內部人交易
  - 國會交易
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
DST = "data/美股_全景儀表板.xlsx"
WATCHLIST_FILE = "data/watchlist_us.txt"
WORKERS = int(os.environ.get("WORKERS", "6"))


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip().upper() for t in env.replace(",", " ").split() if t.strip()]
        toks = [t for t in toks if t and not t.startswith("#")]
        if toks: return list(dict.fromkeys(toks))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip().upper() for t in line.split() if t.strip())
        if toks: return list(dict.fromkeys(toks))
    return "NVDA AVGO TSM META GOOG MSFT".split()


def get(endpoint, **params):
    params["apikey"] = KEY
    for attempt in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429: time.sleep(2 * (attempt+1)); continue
        if r.status_code != 200: return None
        try: return r.json()
        except: return None
    return None


def first(d):
    """list[0] or dict 本身, None 防呆"""
    if d is None: return None
    if isinstance(d, list): return d[0] if d else None
    return d


def fetch_one(sym):
    """對一檔抓所有需要的端點, 回傳 dict"""
    try:
        out = {"代號": sym}

        # ─── 1. 報價 ───
        q = first(get("quote", symbol=sym)) or {}
        price = q.get("price")
        out["當前股價"] = price
        out["市值(億美)"] = round(q.get("marketCap", 0) / 1e8, 0) if q.get("marketCap") else None
        out["52w高"] = q.get("yearHigh")
        out["52w低"] = q.get("yearLow")

        # ─── 2. DCF 估值雷達 ───
        dcf = first(get("discounted-cash-flow", symbol=sym)) or {}
        lev = first(get("levered-discounted-cash-flow", symbol=sym)) or {}
        dcf_val = dcf.get("dcf")
        lev_val = lev.get("dcf")
        out["DCF估值"] = round(dcf_val, 1) if dcf_val else None
        out["LeveredDCF"] = round(lev_val, 1) if lev_val else None
        if price and dcf_val and price > 0:
            out["DCF差%"] = round((dcf_val / price - 1) * 100, 1)
        if price and lev_val and price > 0:
            out["LDCF差%"] = round((lev_val / price - 1) * 100, 1)

        # ─── 3. 體質健診 ───
        sc = first(get("financial-scores", symbol=sym)) or {}
        out["AltmanZ"] = round(sc.get("altmanZScore", 0), 2) if sc.get("altmanZScore") is not None else None
        out["Piotroski"] = sc.get("piotroskiScore")

        # ─── 4. Owner Earnings ───
        oe = first(get("owner-earnings", symbol=sym, limit=1)) or {}
        oe_ps = oe.get("ownersEarningsPerShare")
        oe_total = oe.get("ownersEarnings")
        out["OE/股"] = round(oe_ps, 2) if oe_ps else None
        out["OE(億)"] = round(oe_total / 1e8, 1) if oe_total else None
        # Owner Earnings Yield = OE/股 / 股價
        if price and oe_ps and price > 0:
            out["OE殖利率%"] = round(oe_ps / price * 100 * 4, 1)  # × 4 (季 → 年化)

        # ─── 5. 產品結構 (取最新一年) ───
        ps = first(get("revenue-product-segmentation", symbol=sym,
                       period="annual", limit=1)) or {}
        ps_data = ps.get("data", {}) if isinstance(ps, dict) else {}
        if ps_data:
            total = sum(v for v in ps_data.values() if isinstance(v, (int, float)) and v)
            if total > 0:
                # 取最大 3 個產品 + 其占比
                sorted_p = sorted(ps_data.items(), key=lambda x: x[1] or 0, reverse=True)
                tops = sorted_p[:3]
                out["主產品1"] = f"{tops[0][0]} {round(tops[0][1]/total*100,0):.0f}%" if len(tops)>=1 else None
                out["主產品2"] = f"{tops[1][0]} {round(tops[1][1]/total*100,0):.0f}%" if len(tops)>=2 else None
                out["主產品3"] = f"{tops[2][0]} {round(tops[2][1]/total*100,0):.0f}%" if len(tops)>=3 else None
                out["__product_data"] = ps_data
                out["__product_year"] = ps.get("fiscalYear")

        # ─── 6. 地理結構 ───
        gs = first(get("revenue-geographic-segmentation", symbol=sym,
                       period="annual", limit=1)) or {}
        gs_data = gs.get("data", {}) if isinstance(gs, dict) else {}
        if gs_data:
            total = sum(v for v in gs_data.values() if isinstance(v, (int, float)) and v)
            if total > 0:
                sorted_g = sorted(gs_data.items(), key=lambda x: x[1] or 0, reverse=True)
                tops = sorted_g[:3]
                out["主地區1"] = f"{tops[0][0][:12]} {round(tops[0][1]/total*100,0):.0f}%" if len(tops)>=1 else None
                out["主地區2"] = f"{tops[1][0][:12]} {round(tops[1][1]/total*100,0):.0f}%" if len(tops)>=2 else None
                # 中國 / 台灣曝險
                china_keys = [k for k in gs_data if "CHINA" in k.upper() or "中國" in k]
                taiwan_keys = [k for k in gs_data if "TAIWAN" in k.upper() or "台灣" in k]
                china_rev = sum(gs_data[k] or 0 for k in china_keys)
                tw_rev = sum(gs_data[k] or 0 for k in taiwan_keys)
                if china_rev: out["中國營收%"] = round(china_rev / total * 100, 0)
                if tw_rev: out["台灣營收%"] = round(tw_rev / total * 100, 0)
                out["__geo_data"] = gs_data
                out["__geo_year"] = gs.get("fiscalYear")

        # ─── 7. 內部人交易統計(最新季)===
        ins_stats = get("insider-trading/statistics", symbol=sym) or []
        latest = ins_stats[0] if isinstance(ins_stats, list) and ins_stats else {}
        if latest:
            out["內部人季度"] = f"{latest.get('year')}Q{latest.get('quarter')}"
            out["內部人買筆"] = latest.get("acquiredTransactions")
            out["內部人賣筆"] = latest.get("disposedTransactions")
            out["內部人買量"] = latest.get("totalAcquired")
            out["內部人賣量"] = latest.get("totalDisposed")
            ratio = latest.get("acquiredDisposedRatio")
            out["內部人買賣比"] = round(ratio, 3) if ratio is not None else None

        # ─── 8. 國會交易 (最新 5 筆) ===
        sen = get("senate-trades", symbol=sym) or []
        hou = get("house-trades", symbol=sym) or []
        sen_list = sen[:5] if isinstance(sen, list) else []
        hou_list = hou[:5] if isinstance(hou, list) else []
        out["__senate"] = sen_list
        out["__house"] = hou_list
        # 過去 90 天國會交易筆數
        from datetime import datetime, timedelta
        cutoff = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
        sen_recent = [t for t in (sen if isinstance(sen, list) else [])
                      if t.get("transactionDate", "") >= cutoff]
        hou_recent = [t for t in (hou if isinstance(hou, list) else [])
                      if t.get("transactionDate", "") >= cutoff]
        out["國會90d買"] = sum(1 for t in sen_recent + hou_recent if t.get("type", "").lower() in ("purchase","buy"))
        out["國會90d賣"] = sum(1 for t in sen_recent + hou_recent if t.get("type", "").lower() in ("sale","sell"))

        return sym, out
    except Exception as e:
        return sym, {"代號": sym, "__error": str(e)}


def composite_signal(row):
    """綜合訊號: DCF + 體質 + 內部人"""
    score = 0; tags = []
    # DCF
    d = row.get("DCF差%")
    if d is not None:
        if d > 30: score += 2; tags.append("💎深度低估")
        elif d > 10: score += 1; tags.append("🟢低估")
        elif d < -30: score -= 2; tags.append("🔴深度高估")
        elif d < -10: score -= 1; tags.append("🟠高估")
    # Altman Z
    z = row.get("AltmanZ")
    if z is not None:
        if z >= 3: score += 1
        elif z < 1.8: score -= 2; tags.append("💀破產風險")
    # Piotroski
    p = row.get("Piotroski")
    if p is not None:
        if p >= 8: score += 1; tags.append("🟢體質強")
        elif p <= 3: score -= 1; tags.append("🟠體質弱")
    # 內部人
    r = row.get("內部人買賣比")
    if r is not None:
        if r >= 1: score += 1; tags.append("👤內部人淨買")
        elif r < 0.1: score -= 1; tags.append("⚠️內部人大賣")
    # 國會
    g_buy = row.get("國會90d買", 0) or 0
    g_sell = row.get("國會90d賣", 0) or 0
    if g_buy > g_sell + 2: tags.append("🏛️國會淨買")
    elif g_sell > g_buy + 2: tags.append("🏛️國會淨賣")
    return score, " ".join(tags) if tags else "—"


def main():
    if not KEY: print("⚠️ 未設 FMP_API_KEY"); return
    codes = load_watchlist()
    print(f"美股全景儀表板 — {len(codes)} 檔 (平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sym, data = fut.result()
            if data: results[sym] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # 體檢總表 merge
    base = pd.DataFrame()
    try:
        h = pd.read_excel("data/美股體檢總表.xlsx", sheet_name="體檢總表")
        h["代號"] = h["代號"].astype(str)
        base = h[["代號","名稱","產業","評等","品質總分"]]
    except Exception as e:
        print(f"⚠️ 讀體檢總表失敗 {e}")

    # 組總覽
    rows = []
    for sym in codes:
        r = results.get(sym, {})
        if not r: continue
        sc, tags = composite_signal(r)
        r2 = {k: v for k, v in r.items() if not k.startswith("__")}
        r2["綜合分"] = sc
        r2["訊號"] = tags
        rows.append(r2)

    df = pd.DataFrame(rows)
    if not base.empty:
        df = df.merge(base, on="代號", how="left")
        front = ["代號","名稱","產業","評等","品質總分","當前股價",
                 "DCF估值","DCF差%","LeveredDCF","LDCF差%",
                 "AltmanZ","Piotroski","OE/股","OE殖利率%",
                 "主產品1","主產品2","主產品3","中國營收%","台灣營收%",
                 "內部人買賣比","國會90d買","國會90d賣","綜合分","訊號"]
        rest = [c for c in df.columns if c not in front]
        df = df[[c for c in front if c in df.columns] + rest]
    df = df.sort_values("綜合分", ascending=False)

    # ─── 個別模組分頁 ───
    # DCF
    dcf_sheet = df[[c for c in ["代號","名稱","評等","當前股價","DCF估值","LeveredDCF","DCF差%","LDCF差%","52w高","52w低"] if c in df.columns]].copy()
    dcf_sheet = dcf_sheet.sort_values("DCF差%", ascending=False, na_position="last")

    # 體質
    score_sheet = df[[c for c in ["代號","名稱","評等","AltmanZ","Piotroski","OE/股","OE(億)","OE殖利率%","訊號"] if c in df.columns]].copy()
    score_sheet = score_sheet.sort_values("Piotroski", ascending=False, na_position="last")

    # 產品(展開)
    prod_rows = []
    for sym, r in results.items():
        pd_data = r.get("__product_data")
        if not pd_data: continue
        year = r.get("__product_year")
        total = sum(v for v in pd_data.values() if isinstance(v, (int,float)) and v)
        for prod, rev in sorted(pd_data.items(), key=lambda x: x[1] or 0, reverse=True):
            if not rev: continue
            prod_rows.append({
                "代號": sym, "年": year, "產品": prod,
                "營收(百萬)": round(rev/1e6, 0),
                "占比%": round(rev/total*100, 1) if total>0 else None
            })
    prod_sheet = pd.DataFrame(prod_rows)

    # 地理
    geo_rows = []
    for sym, r in results.items():
        gd = r.get("__geo_data")
        if not gd: continue
        year = r.get("__geo_year")
        total = sum(v for v in gd.values() if isinstance(v, (int,float)) and v)
        for region, rev in sorted(gd.items(), key=lambda x: x[1] or 0, reverse=True):
            if not rev: continue
            geo_rows.append({
                "代號": sym, "年": year, "地區": region,
                "營收(百萬)": round(rev/1e6, 0),
                "占比%": round(rev/total*100, 1) if total>0 else None
            })
    geo_sheet = pd.DataFrame(geo_rows)

    # 內部人
    ins_sheet = df[[c for c in ["代號","名稱","內部人季度","內部人買筆","內部人賣筆","內部人買量","內部人賣量","內部人買賣比"] if c in df.columns]].copy()

    # 國會交易(展開)
    cong_rows = []
    for sym, r in results.items():
        for tx in (r.get("__senate") or [])[:5]:
            cong_rows.append({
                "代號": sym, "院": "參議院",
                "議員": f"{tx.get('firstName','')} {tx.get('lastName','')}".strip(),
                "日期": tx.get("transactionDate"),
                "類型": tx.get("type"),
                "金額": tx.get("amount"),
            })
        for tx in (r.get("__house") or [])[:5]:
            cong_rows.append({
                "代號": sym, "院": "眾議院",
                "議員": f"{tx.get('firstName','')} {tx.get('lastName','')}".strip(),
                "日期": tx.get("transactionDate"),
                "類型": tx.get("type"),
                "金額": tx.get("amount"),
            })
    cong_sheet = pd.DataFrame(cong_rows)
    if not cong_sheet.empty:
        cong_sheet = cong_sheet.sort_values("日期", ascending=False)

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="總覽", index=False)
        dcf_sheet.to_excel(xw, sheet_name="DCF估值", index=False)
        score_sheet.to_excel(xw, sheet_name="體質健診", index=False)
        if not prod_sheet.empty: prod_sheet.to_excel(xw, sheet_name="產品結構", index=False)
        if not geo_sheet.empty: geo_sheet.to_excel(xw, sheet_name="地理結構", index=False)
        ins_sheet.to_excel(xw, sheet_name="內部人交易", index=False)
        if not cong_sheet.empty: cong_sheet.to_excel(xw, sheet_name="國會交易", index=False)

    print(f"\n→ 已輸出 {DST}")
    print(f"分頁: 總覽 / DCF估值 / 體質健診 / 產品結構 / 地理結構 / 內部人交易 / 國會交易")
    print(f"\n=== 綜合分 TOP 15 ===")
    show_cols = [c for c in ["代號","名稱","評等","DCF差%","AltmanZ","Piotroski","內部人買賣比","綜合分","訊號"] if c in df.columns]
    print(df[show_cols].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
