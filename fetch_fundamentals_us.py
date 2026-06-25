# -*- coding: utf-8 -*-
"""
美股 財報 + 估值 + 體檢 (FMP 版)
=======================================================================
讀 tickers_us.txt,用 Financial Modeling Prep (FMP) 抓財報＋估值＋逐季毛利＋歷史 PE/PB,
套用與台股對等的「找好公司 ①~⑩ 框架 + 循環股例外」打分,輸出 data/美股體檢總表.xlsx。

★ 為何改用 FMP(取代 yfinance):
   - yfinance 季報只有 ~4 季、無歷史 PE/PB 序列、無 ROIC、欄位常缺
   - FMP 提供逐季財報、5+ 年歷史比率、ROIC/ROE,跟台股 FinMind 資料齊整度對等

★ 需設 secret: FMP_API_KEY (https://site.financialmodelingprep.com/)
   免費 250/day → 258 檔 × 5 endpoint = 1290,要 Starter 方案(300/min)。
   無 key 則退回:跳過,提示使用者去設(免費註冊即得 key)。

API endpoints:
  /api/v3/income-statement/{sym}?period=annual    年度損益(取 5 年算 CAGR)
  /api/v3/income-statement/{sym}?period=quarter   逐季損益(算逐季毛利+季 YoY)
  /api/v3/cash-flow-statement/{sym}?period=annual 年度現金流(算含金量)
  /api/v3/key-metrics-ttm/{sym}                   TTM:ROIC/ROE/EPS/含金量/股息
  /api/v3/ratios/{sym}?period=annual              5 年 PER/PB → 算位階

每檔算:
  營收/EPS 年序列 → 5y/3y CAGR、ROE、ROIC、毛利率、淨利率、含金量、PER、PBR
  PE/PB 位階(5 年百分位,跟台股對等)
  逐季毛利率(8 季,供 L2 拐點掃描算毛利拐頭)
品質總分(0~100) + 評等 A/B/C/D + 估值標籤(用真位階) + 循環旗 + 主要漏洞
輸出分頁:體檢總表 / A級好公司 / A級+好價格 / 循環股 / 逐季毛利率
"""
import os, time
import numpy as np
import pandas as pd
import requests
from forward_pe import forward_metrics      # 未來估值(Forward PE/PEG)單一真理來源,與台股同口徑

BASE  = "https://financialmodelingprep.com/stable"     # 2025/8/31 後 v3 變 legacy,新用戶須用 stable

WATCH = os.environ.get("US_WATCH_FILE", "tickers_us.txt")    # 可指定 tickers_us_core.txt
CORE  = "tickers_us_core.txt"          # 持倉/曾持有,永遠優先抓
OUT   = "data/美股體檢總表.xlsx"
KEY   = os.environ.get("FMP_API_KEY", "")

RATE_SLEEP = 0.25                  # 每檔間隔(starter ~300/分)
MAX_FETCH  = int(os.environ.get("US_MAX_FETCH", "45"))   # 每輪上限(45×5≈225 calls < 免費250/day)
FRESH_DAYS = int(os.environ.get("US_FRESH_DAYS", "30"))  # 體檢表內 N 天內的不重抓(增量省額度)


def load_watch():
    """合併 CORE(持倉優先)+ WATCH(全名單),去重保持順序。"""
    out = []
    for path in (CORE, WATCH):
        if not os.path.exists(path):
            continue
        for line in open(path, encoding="utf-8"):
            line = line.split("#", 1)[0]
            for tok in line.replace(",", " ").split():
                out.append(tok.strip().upper())
    return list(dict.fromkeys([t for t in out if t]))


def load_existing():
    """讀現有體檢總表 → (rows_dict, q_gm_dict)。供增量:已抓過的不重抓。"""
    if not os.path.exists(OUT):
        return {}, {}
    try:
        df = pd.read_excel(OUT, "體檢總表"); df["代號"] = df["代號"].astype(str)
        rows = {r["代號"]: dict(r) for _, r in df.iterrows()}
    except Exception:
        rows = {}
    qg = {}
    try:
        q = pd.read_excel(OUT, "逐季毛利率", index_col=0)
        for k in q.index:
            qg[str(k)] = q.loc[k].dropna().to_dict()
    except Exception:
        pass
    return rows, qg


def get(endpoint, **params):
    """stable 端點:symbol 與 period/limit 都用 query parameter。"""
    if not KEY:
        raise RuntimeError("未設 FMP_API_KEY")
    params["apikey"] = KEY
    r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=20)
    if r.status_code == 429:
        raise RuntimeError("rate-limit")
    r.raise_for_status()
    return r.json()


def cagr(v, n):
    if len(v) >= n + 1 and v[-(n+1)] > 0 and v[-1] > 0:
        return (v[-1] / v[-(n+1)]) ** (1 / n) - 1
    return np.nan


def is_cyclical(v):
    if len(v) < 3:
        return False
    if min(v) <= 0:
        return True
    return any(v[i] < v[i-1] * 0.8 for i in range(1, len(v)))


def historical_pctl(values, current):
    """current 值在歷史 values 中的百分位 (0~100,越低越便宜)。"""
    v = pd.Series([x for x in values if pd.notna(x) and x > 0])
    if v.empty or pd.isna(current) or current <= 0:
        return np.nan
    return round(float((v <= current).mean() * 100))


def fetch_one(sym):
    inc_a  = get("income-statement",      symbol=sym, period="annual",  limit=6)
    inc_q  = get("income-statement",      symbol=sym, period="quarter", limit=8)
    cf_a   = get("cash-flow-statement",   symbol=sym, period="annual",  limit=6)
    ttm    = get("key-metrics-ttm",       symbol=sym)                     # ROIC/ROE/含金量等品質
    km_a   = get("key-metrics",           symbol=sym, period="annual",  limit=5)  # 歷史 PER/PB 用這個
    ratios = get("ratios",                symbol=sym, period="annual",  limit=5)
    prof   = get("profile",               symbol=sym)

    # 年序列(由舊到新)
    inc_a = list(reversed(inc_a))
    cf_a  = list(reversed(cf_a))
    ratios = list(reversed(ratios))

    rev = [x.get("revenue") for x in inc_a if x.get("revenue") is not None]
    gp  = [x.get("grossProfit") for x in inc_a if x.get("grossProfit") is not None]
    ni  = [x.get("netIncome") for x in inc_a if x.get("netIncome") is not None]
    eps = [x.get("epsdiluted") or x.get("eps") for x in inc_a]
    eps = [x for x in eps if x is not None]
    ocf = [x.get("operatingCashFlow") for x in cf_a if x.get("operatingCashFlow") is not None]

    # 逐季毛利(8 季,由舊到新)
    inc_q = list(reversed(inc_q))
    q_gm = []
    for x in inc_q:
        rv, g = x.get("revenue"), x.get("grossProfit")
        if rv and g:
            q_gm.append((x.get("date", "")[:10], round(g/rv*100, 2)))
    # 近四季 EPS(TTM)= 最近 4 季 epsdiluted 加總(台股口徑:自算 PE,不信廠商年度EPS/比率)
    q_eps = [(x.get("epsdiluted") or x.get("eps")) for x in inc_q]
    q_eps = [e for e in q_eps if e is not None]
    ttm_eps = sum(q_eps[-4:]) if len(q_eps) >= 4 else (eps[-1] if eps else None)

    ttm0 = ttm[0] if ttm else {}
    prof0 = prof[0] if prof else {}

    def pick(d, *keys):
        """從 dict 依序試多個鍵名,回第一個非None。FMP stable 端點鍵名常變。"""
        for k in keys:
            v = d.get(k)
            if v is not None: return v
        return None

    # 5 年歷史 PE/PB → 位階(改用 key-metrics annual,stable 把 ratios 拆了沒這個)
    km_a_sorted = list(reversed(km_a))
    hist_pe = [pick(r, "peRatio", "priceEarningsRatio", "priceToEarningsRatio") for r in km_a_sorted]
    hist_pb = [pick(r, "pbRatio", "priceToBookRatio", "priceToBookValueRatio") for r in km_a_sorted]
    # 5 年歷史 ROE(供動態惡化判斷)— 取近 5 年平均
    hist_roe = [pick(r, "roe", "returnOnEquity") for r in km_a_sorted]
    roe_avg = np.mean([x for x in hist_roe if x is not None]) * 100 if any(x is not None for x in hist_roe) else None
    # 流動比(FMP currentRatio TTM)
    _cur = pick(ttm0, "currentRatioTTM", "currentRatio")
    # TTM PER/PBR:stable 的 key-metrics-ttm 不含,改自算(profile.price ÷ EPS, 跟台股同款)
    price = prof0.get("price") or prof0.get("regularMarketPrice")
    # 自算 PER = 收盤價 ÷ 近四季 TTM EPS(台股同口徑;比年度EPS即時、避免一次性/時點錯位)
    cur_pe = (price / ttm_eps) if (price and ttm_eps and ttm_eps > 0) else None
    # ADR/EPS換算失真 sanity gate:正獲利公司 PER<6 幾乎必為 EPS 基準錯亂(ADR比例/幣別),
    # 標 None 不可信(PDD/TSM/BABA/CMCSA 等都 <5,真實最低合理 PER 約 9-10,中間有空檔)
    if cur_pe is not None and cur_pe < 6:
        cur_pe = None
    # PBR 自算需淨值,先試 ttm 的 bookValuePerShare,沒有則跳過
    bvps = pick(ttm0, "bookValuePerShareTTM", "tangibleBookValuePerShareTTM")
    cur_pb = (price / bvps) if (price and bvps and bvps > 0) else None
    # TTM ROE/ROIC(key-metrics-ttm 確認有)
    _roe = pick(ttm0, "returnOnEquityTTM", "roeTTM")
    _roic = pick(ttm0, "returnOnInvestedCapitalTTM", "roicTTM", "returnOnCapitalEmployedTTM")
    _peg = None              # stable 無 TTM PEG,留空(可未來自算 PER/EPS3y)
    _dy = pick(prof0, "lastDividend", "dividendYield")  # profile 有 dividend
    if _dy and price and _dy > 1:  # 若是絕對股息金額,換算殖利率
        _dy = _dy / price
    _debt = pick(ttm0, "debtToAssetsTTM", "totalDebtToAssetsTTM",
                       "debtRatioTTM", "totalDebtToTotalAssetsTTM")
    # 退路:若 TTM 鍵名變動抓不到,從最新 annual key-metrics 取
    if _debt is None and km_a:
        latest_km = km_a[0]
        _debt = pick(latest_km, "debtToAssets", "totalDebtToAssets",
                                "debtRatio", "totalDebtToTotalAssets")

    rev_l = rev[-1] if rev else None
    ni_l  = ni[-1] if ni else None
    eps_l = ttm_eps                      # forward 用近四季 TTM EPS(非年度),與自算 PER 同口徑

    return {
        "代號": sym,
        "名稱": prof0.get("companyName", sym),
        "產業": prof0.get("sector", ""),
        "收盤": prof0.get("price"),
        "市值(億美)": round(prof0.get("mktCap", 0) / 1e8, 1) if prof0.get("mktCap") else np.nan,
        "營收CAGR%": round(cagr(rev, min(4, len(rev)-1)) * 100, 1) if len(rev) >= 2 else np.nan,
        "毛利率%": round(gp[-1] / rev_l * 100, 1) if gp and rev_l else np.nan,
        "營益率%": round(inc_a[-1].get("operatingIncomeRatio", 0) * 100, 1) if inc_a else np.nan,
        "淨利率%": round(ni_l / rev_l * 100, 1) if ni_l and rev_l else np.nan,
        "EPS5y%": round(cagr(eps, min(4, len(eps)-1)) * 100, 1) if len(eps) >= 2 else np.nan,
        "EPS3y%": round(cagr(eps, 3) * 100, 1) if len(eps) >= 4 else np.nan,
        "ROE%":   round(_roe * 100, 1) if _roe else np.nan,
        "ROIC%":  round(_roic * 100, 1) if _roic else np.nan,
        "含金量": round(ocf[-1] / ni_l, 2) if ocf and ni_l else np.nan,
        "負債比%": round(_debt * 100, 1) if _debt else np.nan,
        "流動比": round(_cur, 2) if _cur else np.nan,
        "ROE5年均%": round(roe_avg, 1) if roe_avg else np.nan,
        "PER":  round(cur_pe, 1) if cur_pe else np.nan,
        "PE位階": historical_pctl(hist_pe, cur_pe),
        "PBR":  round(cur_pb, 2) if cur_pb else np.nan,
        "PBR位階": historical_pctl(hist_pb, cur_pb),
        "PEG":  round(_peg, 2) if _peg else np.nan,
        "殖利率%": round(_dy * 100, 2) if _dy else 0,
        "_eps_l": eps_l,
        "_eps_series": eps,
        "_q_gm": q_gm,
    }


def valuation_tag(pe_pos, pb_pos):
    """用 5 年位階(跟台股對等)分級。"""
    xs = [x for x in (pe_pos, pb_pos) if pd.notna(x)]
    if not xs: return "—"
    m = np.mean(xs)
    if m <= 30: return "🟢便宜"
    if m <= 55: return "🟡合理"
    if m <= 80: return "🟠偏貴"
    return "🔴過熱"


def grade(r):
    s, leak = 0.0, []
    e5, e3 = r["EPS5y%"], r["EPS3y%"]
    if pd.notna(e5) and pd.notna(e3):
        if e5 >= 10 and e3 >= 10: p = 20
        elif e5 > 0 and e3 > 0:   p = 12
        elif (e5 < 0) ^ (e3 < 0): p = 4; leak.append("EPS單期衰退")
        else:                     p = 0; leak.append("EPS連年衰退")
    else: p = 0; leak.append("EPS資料不足")
    s += p
    g = r["含金量"]
    if pd.isna(g):    p = 0; leak.append("無現金資料")
    elif g >= 1.2:    p = 20
    elif g >= 1.0:    p = 16
    elif g >= 0.7:    p = 10
    elif g >= 0.5:    p = 4;  leak.append(f"含金量{g}弱")
    else:             p = 0;  leak.append(f"含金量{g}差")
    s += p
    roe = r["ROE%"]
    if pd.isna(roe):  p = 0
    elif roe >= 20:   p = 14
    elif roe >= 15:   p = 10
    elif roe >= 12:   p = 6
    elif roe >= 8:    p = 3
    else:             p = 0; leak.append(f"ROE{roe}低")
    s += p
    gm = r["毛利率%"]
    p = 10 if (pd.notna(gm) and gm >= 40) else 6 if (pd.notna(gm) and gm >= 25) else 2 if pd.notna(gm) else 0
    s += p
    nm = r["淨利率%"]
    p = 10 if (pd.notna(nm) and nm >= 15) else 6 if (pd.notna(nm) and nm >= 8) else 2 if pd.notna(nm) else 0
    s += p
    rc = r["營收CAGR%"]
    if pd.isna(rc):   p = 0
    elif rc >= 10:    p = 8
    elif rc >= 0:     p = 5
    else:             p = 0; leak.append(f"營收萎縮{rc}%")
    s += p
    if pd.notna(e5) and pd.notna(rc) and rc > 0:
        p = 8 if e5 >= rc else 4 if e5 >= 0.5*rc else 0
        if p == 0: leak.append("EPS落後營收(稀釋/毛利漏)")
    else:
        p = 4 if (pd.notna(e5) and e5 > 0) else 0
    s += p
    # ⑪ 動態惡化扣分(同台股,最高扣 -15)
    penalty = 0
    roe_cur = r.get("ROE%"); roe_avg = r.get("ROE5年均%")
    if pd.notna(roe_cur) and pd.notna(roe_avg) and roe_avg >= 15 and roe_cur < roe_avg * 0.67:
        penalty += 10
        leak.append(f"⚠️ROE滑落({roe_cur:.0f}<5年均{roe_avg:.0f}×67%)")
    dr = r.get("負債比%"); cr_us = r.get("流動比")
    # US 流動比是小數(1.5 = 150%),轉成百分比比較;若無 currentRatio 退而求其次只看負債
    cur_pct = cr_us * 100 if pd.notna(cr_us) else None
    if pd.notna(dr) and dr > 70 and pd.notna(cur_pct) and cur_pct < 100:
        penalty += 10
        leak.append(f"⚠️短期償債警報(負債{dr:.0f}+流動{cur_pct:.0f})")
    elif pd.notna(dr) and dr > 80:
        penalty += 5
        leak.append(f"⚠️高槓桿(負債{dr:.0f}%)")
    penalty = min(penalty, 15)
    s -= penalty
    return round(s, 1), "、".join(leak[:3])


def main():
    if not KEY:
        print("⚠️ 未設 FMP_API_KEY 環境變數,腳本中止。")
        print("   到 https://site.financialmodelingprep.com 免費註冊取得 API key,")
        print("   設 GitHub Secret 名為 FMP_API_KEY,或本地 export FMP_API_KEY=...")
        return
    watch = load_watch()
    existing, q_gm_all = load_existing()          # 增量:讀現有體檢表當快取
    # 待抓 = 名單內但體檢表還沒有的(持倉/曾持有因在 CORE 排前面 → 優先);每輪上限 MAX_FETCH
    todo = [s for s in watch if s not in existing][:MAX_FETCH]
    print(f"美股名單 {len(watch)} 檔 | 已有 {len(existing)} | 本輪補抓 {len(todo)} 檔"
          f"(上限 {MAX_FETCH},約 {len(todo)*5} calls)")
    if not todo:
        print("✓ 名單全部已抓過(增量完成);要強制更新請刪 data/美股體檢總表.xlsx 或調 US_MAX_FETCH")
    rows = list(existing.values())                 # 先保留既有,新抓的 append
    debug_done = False
    for i, sym in enumerate(todo, 1):
        try:
            if not debug_done:                 # 第一檔印 ttm/ratios 鍵名,供校準
                _t = get("key-metrics-ttm", symbol=sym)
                _r = get("ratios", symbol=sym, period="annual", limit=5)
                _km = get("key-metrics", symbol=sym, period="annual", limit=1)
                ttm_keys = list((_t[0] if _t else {}).keys())
                km_keys  = list((_km[0] if _km else {}).keys())
                debt_ttm = [k for k in ttm_keys if "debt" in k.lower() or "leverage" in k.lower()]
                debt_km  = [k for k in km_keys  if "debt" in k.lower() or "leverage" in k.lower()]
                print(f"DEBUG {sym} ttm debt keys:", debt_ttm)
                print(f"DEBUG {sym} km annual debt keys:", debt_km)
                debug_done = True
            r = fetch_one(sym)
            v_eps = r.pop("_eps_series")
            q_gm  = r.pop("_q_gm")
            r_eps_l = r.pop("_eps_l", None)
            cyc = is_cyclical(v_eps)
            sc, leak = grade(r)
            r["品質總分"] = sc
            r["評等"] = "A" if sc >= 80 else "B" if sc >= 65 else "C" if sc >= 50 else "D"
            r["估值"] = valuation_tag(r["PE位階"], r["PBR位階"])
            r["循環股"] = "⚠️循環(看PBR)" if cyc else ""
            # PEG 自算 + Forward PE(與台股同口徑 forward_pe 模組;美股無月營收→保守情境傳 None)
            e3 = r.get("EPS3y%"); e5 = r.get("EPS5y%")
            per = r.get("PER"); close = r.get("收盤"); ttm_eps = r_eps_l
            if per is not None and e3 is not None and not pd.isna(per) and not pd.isna(e3) and e3 > 0:
                r["PEG"] = round(per / e3, 2)
            fwd = forward_metrics(close, ttm_eps, per, e3, e5, None, cyc)
            for k, vv in fwd.items():
                if k == "PEG" and not pd.isna(r.get("PEG", float("nan"))):
                    continue          # PEG 已自算,不覆蓋
                r[k] = vv
            r["主要漏洞"] = leak
            rows.append(r)
            if q_gm: q_gm_all[f"{sym} {r['名稱'][:18]}"] = dict(q_gm)
            print(f"[{i}/{len(todo)}] {sym:6s} 分 {sc} {r['評等']} {r['估值']}")
        except Exception as e:
            print(f"  ! {sym} 失敗:{e}")
        time.sleep(RATE_SLEEP)

    if not rows:
        print("⚠️ 0 檔成功,不產出 Excel(可能是 API key/端點問題,先查 log)")
        return
    df = pd.DataFrame(rows).sort_values("品質總分", ascending=False)
    cols = ["代號", "名稱", "產業", "評等", "品質總分", "EPS5y%", "EPS3y%",
            "ROE%", "ROE5年均%", "ROIC%", "負債比%", "流動比", "含金量",
            "毛利率%", "淨利率%", "營收CAGR%", "PER", "PE位階", "PBR", "PBR位階", "PEG",
            "成長率g%", "預估明年EPS", "ForwardPE", "未來估值",
            "估值", "殖利率%", "市值(億美)", "循環股", "主要漏洞"]
    df = df[[c for c in cols if c in df.columns]]
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="體檢總表", index=False)
        a = df[df["評等"] == "A"]
        a.to_excel(xw, sheet_name="A級好公司", index=False)
        a[a["估值"].isin(["🟢便宜", "🟡合理"])].to_excel(xw, sheet_name="A級+好價格", index=False)
        df[df["循環股"] != ""].to_excel(xw, sheet_name="循環股", index=False)
        if q_gm_all:                                  # 逐季毛利率(供拐點掃描算毛利拐頭)
            qg = pd.DataFrame({k: pd.Series(v) for k, v in q_gm_all.items()}).T
            qg = qg.reindex(columns=sorted(qg.columns))
            qg.to_excel(xw, sheet_name="逐季毛利率")
    print(f"完成 → {OUT}({len(df)} 檔)")


if __name__ == "__main__":
    main()
