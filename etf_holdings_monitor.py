#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主動式 ETF 持股監看  etf_holdings_monitor.py
=====================================================================
跨投信抓「主動式 ETF」每日成分,diff 昨天 → 看經理人在買賣什麼。
主動式 ETF = 真人經理人選股,持股變化是「法人決策」,訊號遠強於被動 ETF。

資料源(本機被 proxy 擋,需在 CI 跑):
  復華 fhtrust : GET /api/assetsExcel/{etf}/{YYYYMMDD}      (Excel)
  統一 ezmoney : GET /ETF/Transaction/PCFExcelNPOI?fundCode={fc}&date={ROC}&specificDate=true

每檔:存日期快照 data/etf_holdings/{code}_{YYYYMMDD}.csv → diff 最近兩日 →
  🟢新增成分 / 🔴剔除 / ⬆️⬇️增減碼。
跨基金:🔥 多檔主動 ETF「同時新增/加碼同一股」= 法人共識(最強訊號)。
"""
import os
import io
import sys
import glob
from datetime import date, timedelta
import requests
import pandas as pd

OUTDIR = "data/etf_holdings"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# 主動式 ETF 監看清單(台股權益型;經理人選股)
FUNDS = [
    {"code": "00991A", "name": "復華未來50",   "house": "fhtrust", "etf": "ETF23"},
    {"code": "00998A", "name": "復華金融股息", "house": "fhtrust", "etf": "ETF24"},
    {"code": "00981A", "name": "統一台股增長", "house": "ezmoney", "fc": "49YTW"},
    {"code": "00403A", "name": "統一升級50",   "house": "ezmoney", "fc": "63YTW"},
]


def roc(d):
    return f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"


def fetch(fund, d):
    """回傳 Excel bytes 或 None。"""
    s = requests.Session(); s.headers["User-Agent"] = UA
    try:
        if fund["house"] == "fhtrust":
            url = f"https://www.fhtrust.com.tw/api/assetsExcel/{fund['etf']}/{d:%Y%m%d}"
            r = s.get(url, timeout=30)
        else:
            url = "https://www.ezmoney.com.tw/ETF/Transaction/PCFExcelNPOI"
            r = s.get(url, params={"fundCode": fund["fc"], "date": roc(d), "specificDate": "true"}, timeout=30)
        if r.status_code == 200 and len(r.content) > 500:
            return r.content
    except Exception as e:
        print(f"  {fund['code']} 抓取失敗:{str(e)[:60]}")
    return None


def parse(content):
    """兩家格式通用:掃描含『代號』的表頭列,標準化為 代號/名稱/股數/權重。"""
    raw = pd.read_excel(io.BytesIO(content), header=None)
    hdr = None
    for i in range(min(30, len(raw))):
        if raw.iloc[i].astype(str).str.contains("代號").any():
            hdr = i; break
    if hdr is None:
        return None
    df = pd.read_excel(io.BytesIO(content), header=hdr)

    def col(keys):
        for c in df.columns:
            if any(k in str(c) for k in keys):
                return c
        return None
    ci, cn, cs, cw = col(["代號"]), col(["名稱"]), col(["股數"]), col(["權重", "比重"])
    if not (ci and cs):
        return None
    out = pd.DataFrame({
        "代號": df[ci].astype(str).str.replace(r"\.0$", "", regex=True).str.strip(),
        "名稱": df[cn].astype(str).str.strip() if cn else "",
        "股數": pd.to_numeric(df[cs].astype(str).str.replace(",", ""), errors="coerce"),
        "權重": pd.to_numeric(df[cw].astype(str).str.replace("%", ""), errors="coerce") if cw else None,
    })
    out = out[out["代號"].str.match(r"^\d{4}$")].dropna(subset=["股數"])
    return out.set_index("代號")


def snapshot_path(code, d):
    return os.path.join(OUTDIR, f"{code}_{d:%Y%m%d}.csv")


def save_snapshot(code, d, df):
    os.makedirs(OUTDIR, exist_ok=True)
    df.reset_index().to_csv(snapshot_path(code, d), index=False, encoding="utf-8-sig")


def prev_snapshot(code, before):
    """讀此 code 的前一個快照(早於 before 的最新一個)。"""
    files = sorted(glob.glob(os.path.join(OUTDIR, f"{code}_*.csv")))
    files = [f for f in files if f < snapshot_path(code, before)]
    if not files:
        return None
    df = pd.read_csv(files[-1], dtype={"代號": str})
    return df.set_index("代號"), os.path.basename(files[-1])


def diff(cur, prev):
    sa, sb = set(prev.index), set(cur.index)
    new = [(s, cur.loc[s, "名稱"], cur.loc[s, "權重"]) for s in sb - sa]
    drop = [(s, prev.loc[s, "名稱"], prev.loc[s, "權重"]) for s in sa - sb]
    chg = []
    for s in sa & sb:
        p0 = prev.loc[s, "股數"]
        d = cur.loc[s, "股數"] - p0
        if not d or not p0:
            continue
        w0 = prev.loc[s, "權重"]
        # 期初基數過小(權重<0.1% 或 股數極微)→ 視為「建倉」,不算百分比
        # (否則 分母趨0 → +84344% 之類的爆表假訊號)
        build = (pd.notna(w0) and w0 < 0.1) or p0 < 1000
        pct = None if build else d / p0 * 100
        chg.append((s, cur.loc[s, "名稱"], pct, cur.loc[s, "權重"], build))
    # 建倉(pct=None)排最前,其餘按變動幅度大→小
    chg.sort(key=lambda x: (1e9 if x[2] is None else abs(x[2])), reverse=True)
    return new, drop, chg


def pull_day(d):
    """抓+存當日所有基金快照,回傳 {code: df}。"""
    out = {}
    for fund in FUNDS:
        content = fetch(fund, d)
        if content is None:
            continue
        cur = parse(content)
        if cur is None or cur.empty:
            continue
        save_snapshot(fund["code"], d, cur)
        out[fund["code"]] = cur
    return out


def cross_fund(day_dfs):
    """跨基金分析:共同持有(共識核心)+ 買賣共識/分歧。"""
    codes = [c for c in day_dfs if not day_dfs[c].empty]
    if len(codes) < 2:
        return
    nm = {f["code"]: f["name"] for f in FUNDS}
    sets = [set(day_dfs[c].index) for c in codes]
    common = set.intersection(*sets)
    ref = day_dfs[codes[0]]
    print("🔥 跨基金『共同持有』(法人共識核心):")
    for s in sorted(common, key=lambda s: -(ref.loc[s, "權重"] if s in ref.index else 0))[:12]:
        ws = " | ".join(f"{nm[c][:2]}{day_dfs[c].loc[s, '權重']:.1f}%" for c in codes if s in day_dfs[c].index)
        nmstr = ref.loc[s, "名稱"] if s in ref.index else s
        print(f"   {s} {str(nmstr)[:5]:5s}  {ws}")
    print(f"   共同 {len(common)}檔 / 共 {len(codes)} 檔基金\n")


def diff_consensus(today):
    """以已存快照,算今日 vs 前一快照的跨基金買賣共識/分歧。"""
    buy, sell = {}, {}
    nm = {f["code"]: f["name"] for f in FUNDS}
    for fund in FUNDS:
        cur_f = snapshot_path(fund["code"], today)
        if not os.path.exists(cur_f):
            continue
        cur = pd.read_csv(cur_f, dtype={"代號": str}).set_index("代號")
        p = prev_snapshot(fund["code"], today)
        if not p:
            continue
        prev, pname = p
        new, drop, chg = diff(cur, prev)
        print(f"[{fund['code']} {fund['name']}] vs {pname}:")
        for s, n, w in new:
            print(f"   🟢新增 {s} {n}"); buy.setdefault(s, []).append(fund["name"])
        for s, n, w in drop:
            print(f"   🔴剔除 {s} {n}"); sell.setdefault(s, []).append(fund["name"])
        for s, n, pct, w, build in chg[:5]:
            if build:
                print(f"   🆕建倉 {s} {n} (現權重{w}%)")
                buy.setdefault(s, []).append(fund["name"])
            else:
                print(f"   {'⬆️加' if pct > 0 else '⬇️減'} {s} {n} {pct:+.0f}%")
                (buy if pct > 0 else sell).setdefault(s, []).append(fund["name"])
        if not (new or drop or chg):
            print("   (無變化)")
        print()
    # 共識(≥2檔同方向)+ 分歧(一買一賣)
    mb = {s: set(f) for s, f in buy.items() if len(set(f)) >= 2}
    ms = {s: set(f) for s, f in sell.items() if len(set(f)) >= 2}
    div = {s: (set(buy.get(s, [])), set(sell.get(s, []))) for s in set(buy) & set(sell)}
    if mb:
        print("🔥🔥 共識加碼(多檔同買):", "、".join(f"{s}({'/'.join(v)})" for s, v in mb.items()))
    if ms:
        print("🔥🔥 共識減碼(多檔同砍):", "、".join(f"{s}({'/'.join(v)})" for s, v in ms.items()))
    if div:
        print("⚖️ 分歧(有人買有人砍):", "、".join(f"{s}(買{','.join(b)}|砍{','.join(se)})" for s, (b, se) in div.items()))
    if not (mb or ms or div):
        print("(無跨基金共識/分歧訊號)")


def week_report():
    """讀所有已存快照,做近一週每檔淨變化 + 期初期末共識。"""
    nm = {f["code"]: f["name"] for f in FUNDS}
    print("\n========== 近期持股趨勢(全部已存快照)==========")
    for fund in FUNDS:
        files = sorted(glob.glob(os.path.join(OUTDIR, f"{fund['code']}_*.csv")))
        if len(files) < 2:
            continue
        a = pd.read_csv(files[0], dtype={"代號": str}).set_index("代號")
        b = pd.read_csv(files[-1], dtype={"代號": str}).set_index("代號")
        d0, d1 = files[0][-12:-4], files[-1][-12:-4]
        print(f"\n━━ {fund['code']} {fund['name']}  {d0}→{d1} ━━")
        new, drop, chg = diff(b, a)
        if new: print("  🟢期間新進:", "、".join(f"{s}{n}" for s, n, w in new))
        if drop: print("  🔴期間剔除:", "、".join(f"{s}{n}" for s, n, w in drop))
        for s, n, pct, w, build in [x for x in chg if x[2] is None or abs(x[2]) >= 10][:8]:
            if build:
                print(f"   🆕建倉 {s}{n} (現權重{w}%)")
            else:
                print(f"   {'⬆️加' if pct > 0 else '⬇️減'} {s}{n} {pct:+.0f}% (現權重{w}%)")


def main():
    backfill = int(os.environ.get("ETF_BACKFILL_DAYS", "0") or "0")
    if backfill > 0:
        # 回補:從 today 往回抓 backfill 天(假日自動跳過),建立一週快照
        end = date.today()
        print(f"=== 回補近 {backfill} 天主動式ETF快照 ===\n")
        for i in range(backfill, -1, -1):
            d = end - timedelta(days=i)
            got = pull_day(d)
            print(f"  {d:%Y/%m/%d}: {len(got)} 檔基金有資料")
        week_report()
        return

    _d = os.environ.get("ETF_DATE", "").strip()
    today = pd.to_datetime(_d, format="%Y%m%d").date() if _d else date.today()
    print(f"=== 主動式 ETF 持股監看 {today:%Y/%m/%d} ===\n")
    day_dfs = pull_day(today)
    if not day_dfs:
        print("今日無資料(假日/未公告)"); return
    diff_consensus(today)
    print()
    cross_fund(day_dfs)


if __name__ == "__main__":
    main()
