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
        d = cur.loc[s, "股數"] - prev.loc[s, "股數"]
        if d and prev.loc[s, "股數"]:
            chg.append((s, cur.loc[s, "名稱"], d / prev.loc[s, "股數"] * 100, cur.loc[s, "權重"]))
    chg.sort(key=lambda x: abs(x[2]), reverse=True)
    return new, drop, chg


def main():
    today = date.today()
    consensus_buy = {}   # 代號 → [基金名] 同時新增/加碼
    print(f"=== 主動式 ETF 持股監看 {today:%Y/%m/%d} ===\n")
    for fund in FUNDS:
        content = fetch(fund, today)
        if content is None:
            print(f"[{fund['code']} {fund['name']}] 今日無資料(假日/未公告/抓取失敗)")
            continue
        cur = parse(content)
        if cur is None or cur.empty:
            print(f"[{fund['code']} {fund['name']}] 解析失敗")
            continue
        save_snapshot(fund["code"], today, cur)
        p = prev_snapshot(fund["code"], today)
        if not p:
            print(f"[{fund['code']} {fund['name']}] {len(cur)}檔(首日快照,無前日可比)")
            continue
        prev, pname = p
        new, drop, chg = diff(cur, prev)
        print(f"[{fund['code']} {fund['name']}] vs {pname}:")
        for s, nm, w in new:
            print(f"   🟢新增 {s} {nm} 權重{w}%"); consensus_buy.setdefault(s, []).append(fund["name"])
        for s, nm, w in drop:
            print(f"   🔴剔除 {s} {nm}")
        for s, nm, pct, w in chg[:6]:
            ar = "⬆️加" if pct > 0 else "⬇️減"
            print(f"   {ar} {s} {nm} {pct:+.1f}% 權重{w}%")
            if pct > 0:
                consensus_buy.setdefault(s, []).append(fund["name"])
        if not (new or drop or chg):
            print("   (無變化)")
        print()

    # 跨基金共識:被多檔同時新增/加碼
    multi = {s: fs for s, fs in consensus_buy.items() if len(set(fs)) >= 2}
    if multi:
        print("🔥🔥 法人共識(多檔主動ETF同時加碼/新增):")
        for s, fs in sorted(multi.items(), key=lambda x: -len(set(x[1]))):
            print(f"   {s} ← {'、'.join(sorted(set(fs)))}")
    else:
        print("(今日無跨基金共識訊號)")


if __name__ == "__main__":
    main()
