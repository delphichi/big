# -*- coding: utf-8 -*-
"""
美股 財報 + 估值 + 體檢 (US Fundamentals & Health Check)
=======================================================================
讀 tickers_us.txt(美股觀察名單),用 yfinance 抓財報,套用與台股同一套
「找好公司 ①~⑩ 框架 + 循環股例外」打分,輸出 data/美股體檢總表.xlsx。

★ 資料來源:yfinance(Yahoo Finance)。GitHub Actions runner 網路無限制可正常抓;
  本地容器若 Yahoo 被 egress 白名單擋(query1/2.finance.yahoo.com),需先放行或改用 CI。

每檔算:
  營收/EPS 年序列(近4年) → 營收CAGR、EPS 5y/近3y CAGR(成長引擎,不是只看營收)
  毛利率/營益率/淨利率(最新年) 、ROE、ROIC(NOPAT/投入資本)
  獲利含金量 = 營業現金流 ÷ 淨利(照妖鏡) 、負債比
  估值:PER、PBR、殖利率、PEG;循環股偵測(EPS 曾負或大起大落 → 看 PBR)
品質總分(0~100,估值另計) + 評等 A/B/C/D + 估值標籤 + 循環旗 + 主要漏洞。
輸出分頁:體檢總表 / A級好公司 / A級+好價格 / 循環股。
"""
import time
import numpy as np
import pandas as pd

WATCH = "tickers_us.txt"
OUT = "data/美股體檢總表.xlsx"


def load_watch():
    out = []
    for line in open(WATCH, encoding="utf-8"):
        line = line.split("#", 1)[0]
        for tok in line.replace(",", " ").split():
            out.append(tok.strip().upper())
    return list(dict.fromkeys([t for t in out if t]))


def _row(df, *names):
    """從 yfinance 財報 DataFrame(index=科目, columns=年/季) 取某列,回傳依時間排序的數列。"""
    if df is None or df.empty:
        return []
    for n in names:
        if n in df.index:
            s = df.loc[n].dropna()
            s = s[sorted(s.index)]          # 由舊到新
            return [float(x) for x in s.values]
    return []


def cagr(v, n):
    """取最後 n+1 個資料點算 n 段 CAGR(需頭尾皆正)。"""
    if len(v) >= n + 1 and v[-(n+1)] > 0 and v[-1] > 0:
        return (v[-1] / v[-(n+1)]) ** (1 / n) - 1
    return np.nan


def is_cyclical(v):
    if len(v) < 3:
        return False
    if min(v) <= 0:
        return True
    return (max(v) / min(v)) > 3


def fetch_one(sym):
    import yfinance as yf
    t = yf.Ticker(sym)
    inc, bs, cf = t.income_stmt, t.balance_sheet, t.cashflow
    info = {}
    try:
        info = t.info or {}
    except Exception:
        info = {}

    rev = _row(inc, "Total Revenue")
    gp = _row(inc, "Gross Profit")
    op = _row(inc, "Operating Income", "Operating Income Or Loss")
    ni = _row(inc, "Net Income", "Net Income Common Stockholders")
    eps = _row(inc, "Diluted EPS", "Basic EPS")
    ocf = _row(cf, "Operating Cash Flow", "Total Cash From Operating Activities")
    ta = _row(bs, "Total Assets")
    tl = _row(bs, "Total Liabilities Net Minority Interest", "Total Liab")
    eq = _row(bs, "Stockholders Equity", "Total Stockholder Equity")
    debt = _row(bs, "Total Debt")

    def last(x):
        return x[-1] if x else np.nan

    rev_l, ni_l = last(rev), last(ni)
    r = {
        "代號": sym,
        "名稱": info.get("shortName", sym),
        "產業": info.get("sector", ""),
        "收盤": info.get("currentPrice") or info.get("regularMarketPrice"),
        "市值(億美)": round(info.get("marketCap", 0) / 1e8, 1) if info.get("marketCap") else np.nan,
        "營收CAGR%": round(cagr(rev, min(3, len(rev)-1)) * 100, 1) if len(rev) >= 2 else np.nan,
        "毛利率%": round(last(gp) / rev_l * 100, 1) if gp and rev_l else (round(info.get("grossMargins", np.nan)*100, 1) if info.get("grossMargins") else np.nan),
        "營益率%": round(last(op) / rev_l * 100, 1) if op and rev_l else (round(info.get("operatingMargins", np.nan)*100, 1) if info.get("operatingMargins") else np.nan),
        "淨利率%": round(ni_l / rev_l * 100, 1) if ni_l and rev_l else (round(info.get("profitMargins", np.nan)*100, 1) if info.get("profitMargins") else np.nan),
        "EPS5y%": round(cagr(eps, min(4, len(eps)-1)) * 100, 1) if len(eps) >= 2 else np.nan,
        "EPS3y%": round(cagr(eps, 3) * 100, 1) if len(eps) >= 4 else (round(cagr(eps, len(eps)-1)*100, 1) if len(eps) >= 2 else np.nan),
        "ROE%": round(info.get("returnOnEquity", np.nan) * 100, 1) if info.get("returnOnEquity") else (round(ni_l / last(eq) * 100, 1) if ni_l and last(eq) else np.nan),
        "含金量": round(last(ocf) / ni_l, 2) if ocf and ni_l and ni_l != 0 else np.nan,
        "負債比%": round(last(tl) / last(ta) * 100, 1) if tl and last(ta) else np.nan,
        "PER": round(info.get("trailingPE", np.nan), 1) if info.get("trailingPE") else np.nan,
        "PBR": round(info.get("priceToBook", np.nan), 2) if info.get("priceToBook") else np.nan,
        "PEG": round(info.get("trailingPegRatio", np.nan), 2) if info.get("trailingPegRatio") else np.nan,
        "殖利率%": round(info.get("dividendYield", 0), 2) if info.get("dividendYield") else 0,
        "_eps_series": eps,
    }
    return r


def valuation_tag(per, pbr):
    pts = []
    if pd.notna(per): pts.append(min(per / 0.4, 100) if per > 0 else 100)  # PER~40→100
    xs = [x for x in pts if pd.notna(x)]
    # 簡化:用 PER 絕對值分級(美股無歷史位階)
    if pd.isna(per) or per <= 0:
        return "—"
    if per <= 15: return "🟢便宜"
    if per <= 25: return "🟡合理"
    if per <= 40: return "🟠偏貴"
    return "🔴過熱"


def grade(r):
    s, leak = 0.0, []
    e5, e3 = r["EPS5y%"], r["EPS3y%"]
    # ⑥EPS成長 20
    if pd.notna(e5) and pd.notna(e3):
        if e5 >= 10 and e3 >= 10: p = 20
        elif e5 > 0 and e3 > 0:   p = 12
        elif (e5 < 0) ^ (e3 < 0): p = 4; leak.append("EPS單期衰退")
        else:                     p = 0; leak.append("EPS連年衰退")
    else:
        p = 0; leak.append("EPS資料不足")
    s += p
    # ⑧含金量 20
    g = r["含金量"]
    if pd.isna(g):    p = 0; leak.append("無現金資料")
    elif g >= 1.2:    p = 20
    elif g >= 1.0:    p = 16
    elif g >= 0.7:    p = 10
    elif g >= 0.5:    p = 4;  leak.append(f"含金量{g}弱")
    else:             p = 0;  leak.append(f"含金量{g}差")
    s += p
    # ⑨ROE 14
    roe = r["ROE%"]
    if pd.isna(roe):  p = 0
    elif roe >= 20:   p = 14
    elif roe >= 15:   p = 10
    elif roe >= 12:   p = 6
    elif roe >= 8:    p = 3
    else:             p = 0; leak.append(f"ROE{roe}低")
    s += p
    # ②毛利率 10 / ④淨利率 10 (絕對值,美股無位階)
    gm = r["毛利率%"]
    p = 10 if (pd.notna(gm) and gm >= 40) else 6 if (pd.notna(gm) and gm >= 25) else 2 if pd.notna(gm) else 0
    s += p
    nm = r["淨利率%"]
    p = 10 if (pd.notna(nm) and nm >= 15) else 6 if (pd.notna(nm) and nm >= 8) else 2 if pd.notna(nm) else 0
    s += p
    # ⑤營收成長 8
    rc = r["營收CAGR%"]
    if pd.isna(rc):   p = 0
    elif rc >= 10:    p = 8
    elif rc >= 0:     p = 5
    else:             p = 0; leak.append(f"營收萎縮{rc}%")
    s += p
    # ⑦EPS不落後營收 8
    if pd.notna(e5) and pd.notna(rc) and rc > 0:
        p = 8 if e5 >= rc else 4 if e5 >= 0.5*rc else 0
        if p == 0: leak.append("EPS落後營收(稀釋/毛利漏)")
    else:
        p = 4 if (pd.notna(e5) and e5 > 0) else 0
    s += p
    return round(s, 1), "、".join(leak[:3])


def main():
    watch = load_watch()
    print(f"美股觀察名單 {len(watch)} 檔,開始抓取...")
    rows = []
    for i, sym in enumerate(watch, 1):
        try:
            r = fetch_one(sym)
            cyc = is_cyclical(r.pop("_eps_series"))
            sc, leak = grade(r)
            r["品質總分"] = sc
            r["評等"] = "A" if sc >= 80 else "B" if sc >= 65 else "C" if sc >= 50 else "D"
            r["估值"] = valuation_tag(r["PER"], r["PBR"])
            r["循環股"] = "⚠️循環(看PBR)" if cyc else ""
            r["主要漏洞"] = leak
            rows.append(r)
            print(f"[{i}/{len(watch)}] {sym} 分 {sc} {r['評等']}")
        except Exception as e:
            print(f"  ! {sym} 失敗:{e}")
        time.sleep(0.3)

    df = pd.DataFrame(rows).sort_values("品質總分", ascending=False)
    cols = ["代號", "名稱", "產業", "評等", "品質總分", "EPS5y%", "EPS3y%", "ROE%", "含金量",
            "毛利率%", "淨利率%", "營收CAGR%", "PER", "PBR", "PEG", "估值", "殖利率%",
            "市值(億美)", "循環股", "主要漏洞"]
    df = df[[c for c in cols if c in df.columns]]
    import os
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="體檢總表", index=False)
        a = df[df["評等"] == "A"]
        a.to_excel(xw, sheet_name="A級好公司", index=False)
        a[a["估值"].isin(["🟢便宜", "🟡合理"])].to_excel(xw, sheet_name="A級+好價格", index=False)
        df[df["循環股"] != ""].to_excel(xw, sheet_name="循環股", index=False)
    print(f"完成 → {OUT}({len(df)} 檔)")


if __name__ == "__main__":
    main()
