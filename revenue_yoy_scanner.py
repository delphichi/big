# -*- coding: utf-8 -*-
"""
月營收年增掃描器 (Monthly Revenue YoY Scanner)
================================================
功能:給一份股票代號清單,逐檔檢查「最近 3 年(36 個月)每月營收,是否較去年同期成長」。
資料來源:FinMind  taiwan_stock_month_revenue(每檔只打 1 次 API,很輕量,適合大量掃描)。
輸出   :data/月營收年增掃描.xlsx(單一摘要表,依成長一致性排序)。

每檔算出:
  - 近36月成長月數 / 比較基數(有去年同期可比的月份)、成長比率%
  - 近12月成長月數、近12月樣態(✔/✗ 視覺化)
  - 連續成長月數(從最新月往回算,連幾個月 YoY 為正)
  - 最新月份、最新月年增%、近36月平均年增%
  - 分類:🟢全期強勢 / 🔵多數成長 / 🟡中性 / 🔴轉弱衰退 / —資料不足

清單來源(三選一,優先序由上而下):
  1. SCAN_ALL_LISTED=True  → 自動抓所有上市櫃普通股(忽略下方清單)
  2. tickers.txt / tickers.csv 存在 → 讀檔(txt:一行一個或逗號/空白分隔,# 為註解;
                                       csv:找 代號/stock_id/ticker 欄,否則取第一欄)
  3. 都沒有 → 用下方內建 TICKERS

大量掃描提醒:
  - 免費 FinMind 約 300 次/hr、設 token 約 600 次/hr。掃 ~1800 檔 ≈ 3 小時起跳。
  - 本程式會把每檔結果即時寫進 data/_revenue_scan_progress.csv(斷點續跑用):
    中途被限流/中斷,直接再跑一次會自動跳過已完成的,接著跑。
  - 撞限流會自動等待重試(MAX_RETRY 次)。
"""

import os, sys, time, csv
import pandas as pd
import numpy as np
from datetime import datetime

# ---------- 設定 ----------
TOKEN          = os.environ.get("FINMIND_TOKEN", "")
SCAN_ALL_LISTED = False                       # True 則掃全市場普通股(忽略清單檔/內建清單)
TICKER_FILE_TXT = "tickers.txt"
TICKER_FILE_CSV = "tickers.csv"
OUTPUT          = "data/月營收年增掃描.xlsx"
PROGRESS        = "data/_revenue_scan_progress.csv"

YEARS           = 3                           # 看最近幾年(每月)
LOOKBACK        = YEARS * 12                   # 36 個月
RATE_SLEEP      = 0.3                          # 每檔間隔秒數(降低撞限流機率)
MAX_RETRY       = 4                            # 限流時重試次數
RESUME          = True                         # 是否啟用斷點續跑

# 內建清單(沒有 tickers.txt / tickers.csv 時才用,可自行編輯)
TICKERS = ["2330", "2454", "2412", "1560", "2912", "1476"]

# 分類門檻
RULES = dict(strong_ratio=80,   # 近36月成長比率 ≥ 此值 → 🟢全期強勢
             good_ratio=60,      #                ≥ 此值 → 🔵多數成長
             weak_ratio=40,      #                < 此值 → 🔴轉弱衰退
             min_base=12)        # 可比月份 < 此值 → 視為資料不足


# ---------- 清單載入 ----------
def _parse_txt(path):
    out = []
    for line in open(path, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        for tok in line.replace(",", " ").replace("\t", " ").split():
            tok = tok.strip().upper()
            if tok:
                out.append(tok)
    return out

def _parse_csv(path):
    df = pd.read_csv(path, dtype=str)
    for col in df.columns:
        if str(col).strip().lower() in ("代號", "股票代號", "stock_id", "ticker", "symbol", "code"):
            return [str(x).strip() for x in df[col].dropna()]
    return [str(x).strip() for x in df.iloc[:, 0].dropna()]   # 退而取第一欄

def load_tickers(api=None):
    if SCAN_ALL_LISTED and api is not None:
        info = api.taiwan_stock_info()
        # 只留上市/上櫃的普通股:排除 ETF(00 開頭)、權證/特別股等非 4 碼純數字代號
        m = info[info["type"].isin(["twse", "tpex"])] if "type" in info.columns else info
        ids = sorted(set(str(s) for s in m["stock_id"]
                         if str(s).isdigit() and len(str(s)) == 4 and not str(s).startswith("00")))
        print(f"全市場掃描:取得 {len(ids)} 檔普通股")
        return ids
    if os.path.exists(TICKER_FILE_TXT):
        ids = _parse_txt(TICKER_FILE_TXT); print(f"從 {TICKER_FILE_TXT} 讀到 {len(ids)} 檔")
        return ids
    if os.path.exists(TICKER_FILE_CSV):
        ids = _parse_csv(TICKER_FILE_CSV); print(f"從 {TICKER_FILE_CSV} 讀到 {len(ids)} 檔")
        return ids
    print(f"用內建清單 {len(TICKERS)} 檔")
    return list(TICKERS)


# ---------- FinMind ----------
def make_loader():
    from FinMind.data import DataLoader
    api = DataLoader()
    if TOKEN:
        try:
            api.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("token 登入失敗(改用免費額度):", e)
    return api

def fetch_revenue(api, sid, start):
    """抓單檔月營收;撞限流自動等待重試。回傳 DataFrame 或 None。"""
    for attempt in range(MAX_RETRY):
        try:
            df = api.taiwan_stock_month_revenue(stock_id=sid, start_date=start)
            return df
        except Exception as e:
            msg = str(e).lower()
            if any(k in msg for k in ("limit", "402", "request", "too many", "exceed")):
                wait = 60 * (attempt + 1)
                print(f"  ! 限流,等 {wait}s 後重試 ({attempt+1}/{MAX_RETRY})")
                time.sleep(wait)
            else:
                print(f"  ! {sid} 取數失敗:{e}")
                return None
    print(f"  ! {sid} 重試 {MAX_RETRY} 次仍失敗,跳過")
    return None


# ---------- 分析 ----------
def analyze(df, lookback=LOOKBACK):
    if df is None or df.empty:
        return None
    need = {"revenue", "revenue_year", "revenue_month"}
    if not need.issubset(df.columns):
        return None
    d = df.dropna(subset=["revenue", "revenue_year", "revenue_month"]).copy()
    if d.empty:
        return None
    d["ym"] = d["revenue_year"].astype(int) * 100 + d["revenue_month"].astype(int)
    d = d.drop_duplicates("ym").sort_values("ym")
    rev = dict(zip(d["ym"], d["revenue"].astype(float)))

    # 算每個有「去年同期」可比的月份 YoY
    rows = []   # (ym, revenue, yoy%)
    for ym in sorted(rev):
        y, m = ym // 100, ym % 100
        prev = (y - 1) * 100 + m
        base = rev.get(prev)
        if base and base != 0:
            rows.append((ym, rev[ym], (rev[ym] - base) / abs(base) * 100))
    if not rows:
        return None

    recent = rows[-lookback:]
    n_base = len(recent)
    n_grow = sum(1 for r in recent if r[2] > 0)

    last12 = rows[-12:]
    g12 = sum(1 for r in last12 if r[2] > 0)
    pat12 = "".join("✔" if r[2] > 0 else "✗" for r in last12)

    streak = 0                       # 連續成長月數(從最新往回)
    for r in reversed(rows):
        if r[2] > 0:
            streak += 1
        else:
            break

    latest = rows[-1]
    return {
        "最新月份":          f"{latest[0]//100}/{latest[0]%100:02d}",
        "最新月年增%":       round(latest[2], 1),
        "近36月可比基數":     n_base,
        "近36月成長月數":     n_grow,
        "近36月成長比率%":   round(n_grow / n_base * 100, 1) if n_base else None,
        "近12月成長":        f"{g12}/{len(last12)}",
        "近12月樣態":        pat12,
        "連續成長月數":       streak,
        "近36月平均年增%":   round(float(np.mean([r[2] for r in recent])), 1),
    }

def label(res):
    if not res or res["近36月可比基數"] < RULES["min_base"]:
        return "— 資料不足/新上市"
    ratio = res["近36月成長比率%"]
    if ratio >= RULES["strong_ratio"]:
        return "🟢 全期強勢"
    if ratio >= RULES["good_ratio"]:
        return "🔵 多數成長"
    if ratio < RULES["weak_ratio"]:
        return "🔴 轉弱/衰退"
    return "🟡 中性"


# ---------- 斷點續跑 ----------
def load_done():
    if RESUME and os.path.exists(PROGRESS):
        try:
            df = pd.read_csv(PROGRESS, dtype={"代號": str})
            done = {str(r["代號"]): dict(r) for _, r in df.iterrows()}
            print(f"續跑:已完成 {len(done)} 檔,將跳過")
            return done
        except Exception:
            pass
    return {}

def append_progress(row):
    os.makedirs(os.path.dirname(PROGRESS), exist_ok=True)
    new = not os.path.exists(PROGRESS)
    with open(PROGRESS, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(row.keys()))
        if new:
            w.writeheader()
        w.writerow(row)


# ---------- 主流程 ----------
def main():
    api = make_loader()
    tickers = load_tickers(api)
    start = f"{datetime.now().year - YEARS - 1}-01-01"   # 多墊一年當 YoY 基期

    done = load_done()
    results = list(done.values())
    todo = [t for t in tickers if t not in done]
    print(f"總清單 {len(tickers)} 檔,待掃 {len(todo)} 檔,起始日 {start}\n")

    for i, sid in enumerate(todo, 1):
        df = fetch_revenue(api, sid, start)
        res = analyze(df)
        row = {"代號": sid, "分類": label(res)}
        row.update(res or {})
        results.append(row)
        append_progress(row)
        print(f"[{i}/{len(todo)}] {sid:6s} {row['分類']:14s} "
              f"成長比率 {row.get('近36月成長比率%','-')}% 連續 {row.get('連續成長月數','-')}月")
        time.sleep(RATE_SLEEP)

    # ---- 輸出 Excel(依成長一致性排序)----
    df = pd.DataFrame(results)
    if not df.empty:
        for c in ("近36月成長比率%", "連續成長月數", "近36月成長月數"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.sort_values(["近36月成長比率%", "連續成長月數", "近36月平均年增%"],
                            ascending=False, na_position="last")
        cols = ["代號", "分類", "近36月成長比率%", "近36月成長月數", "近36月可比基數",
                "連續成長月數", "近12月成長", "近12月樣態",
                "最新月份", "最新月年增%", "近36月平均年增%"]
        df = df[[c for c in cols if c in df.columns]]

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="月營收年增掃描", index=False)
    print(f"\n完成 → {OUTPUT}({len(df)} 檔)")
    if not df.empty and "分類" in df.columns:
        print(df["分類"].value_counts().to_string())


if __name__ == "__main__":
    main()
