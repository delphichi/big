# -*- coding: utf-8 -*-
"""
月營收年增掃描器 (Monthly Revenue YoY Scanner)
================================================
功能:給一份股票代號清單,逐檔檢查「最近 10 年(120 個月)每月營收,是否較去年同期成長」,
      並同時並列近10年/近5年/近3年/近1年四個窗的成長比率(看長期 vs 近期)。
資料來源:FinMind  taiwan_stock_month_revenue(每檔只打 1 次 API,很輕量,適合大量掃描)。
輸出   :data/月營收年增掃描.xlsx(摘要表;開啟明細模式時另含逐月年增%明細表)。

每檔算出:
  - 近10年/近5年/近3年/近1年 成長比率%(各窗:YoY 為正的月數 ÷ 該窗實際月數)
  - 近10年/近5年/近3年/近1年 增長月數(計數形式,例 80/120 = 120 個月裡有 80 個月較去年同期增長)
  - 可比月數、連續成長月數(從最新月往回算,連幾個月 YoY 為正)
  - 近12月樣態(✔/✗ 視覺化)、最新月份、最新月年增%、近10年平均年增%
  - 分類:🟢全期強勢 / 🔵多數成長 / 🟡中性 / 🔴轉弱衰退 / —資料不足
    (分類預設用近3年窗,見 CLASS_WINDOW)

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
from datetime import datetime, timedelta

# ---------- 設定 ----------
TOKEN          = os.environ.get("FINMIND_TOKEN", "")
SCAN_ALL_LISTED = False                       # True 則掃全市場普通股(忽略清單檔/內建清單)
TICKER_FILE_TXT = "tickers.txt"
TICKER_FILE_CSV = "tickers.csv"
OUTPUT          = "data/月營收年增掃描.xlsx"
PROGRESS        = "data/_revenue_scan_progress.csv"

YEARS           = 10                          # 看最近幾年(每月)→ 10 年
LOOKBACK        = YEARS * 12                   # 120 個月
RATE_SLEEP      = 0.3                          # 每檔間隔秒數(降低撞限流機率)
MAX_RETRY       = 4                            # 限流時重試次數
RESUME          = True                         # 是否啟用斷點續跑

# 多窗成長比率(月數, 欄名):同時看長期與近期,避免被單一窗誤導
WINDOWS         = [(120, "近10年"), (60, "近5年"), (36, "近3年"), (12, "近1年")]
CLASS_WINDOW    = 36                           # 🟢🔵🟡🔴 分類用哪個窗(月)。
                                               # 預設近3年(門檻才有意義);要嚴格 10 年濾網
                                               # 改成 120,並把下方 RULES 門檻調低
WRITE_MONTHLY_DETAIL = False                   # True 則另輸出「每月營收+年增%」明細表
                                               # ★ 僅適合小清單(如精選名單),全市場會爆量

# 內建清單(沒有 tickers.txt / tickers.csv 時才用,可自行編輯)
TICKERS = ["2330", "2454", "2412", "1560", "2912", "1476"]

# 分類門檻(對應 CLASS_WINDOW;若改用 120 窗,建議 strong=65/good=50/weak=35)
RULES = dict(strong_ratio=80,   # 成長比率 ≥ 此值 → 🟢全期強勢
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

class RateLimited(Exception):
    """FinMind 每小時額度用罄(需等整點重置),與『查無資料』區分開。"""
    pass

def _is_rate_limit(e):
    msg = str(e).lower()
    return any(k in msg for k in ("limit", "402", "429", "request", "too many", "exceed"))

def seconds_to_next_hour(buffer=45):
    """距離下一個整點還有幾秒(額度每小時重置),多加 buffer 秒保險。"""
    now = datetime.now()
    nxt = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    return max(5, int((nxt - now).total_seconds()) + buffer)

def show_usage(api):
    """盡量印出目前 API 用量(不同版本屬性名不一,失敗就略過)。"""
    for attr in ("api_usage", "api_usage_limit"):
        try:
            print(f"  FinMind {attr} = {getattr(api, attr)}")
        except Exception:
            pass

def fetch_revenue(api, sid, start):
    """抓單檔月營收。
    - 瞬間抖動:短重試一次(30s)。
    - 確定是額度用罄:丟出 RateLimited,交給主迴圈『睡到整點再續、且不跳過此檔』。
    - 其他錯誤/真的查無資料:回 None(可正常標記完成)。"""
    for attempt in range(2):
        try:
            return api.taiwan_stock_month_revenue(stock_id=sid, start_date=start)
        except Exception as e:
            if _is_rate_limit(e):
                if attempt == 0:
                    time.sleep(30)
                    continue
                raise RateLimited(str(e))
            print(f"  ! {sid} 取數失敗(非限流):{e}")
            return None
    raise RateLimited("retry exhausted")


# ---------- 分析 ----------
def analyze(df, lookback=LOOKBACK, want_detail=False):
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
    out = {
        "最新月份":   f"{rows[-1][0]//100}/{rows[-1][0]%100:02d}",
        "最新月年增%": round(rows[-1][2], 1),
        "可比月數":   len(recent),
    }
    # 多窗成長比率(各窗:該窗內 YoY 為正的月數 ÷ 該窗實際月數)
    # 同時輸出「增長月數/可比月數」的計數形式(例:80/120),一眼看出 80 個月增長、共 120 個月。
    class_ratio, class_base = None, 0
    for w, name in WINDOWS:
        seg = rows[-w:]
        if seg:
            up = sum(1 for r in seg if r[2] > 0)        # 該窗 YoY 為正(增長)的月數
            total = len(seg)                            # 該窗實際可比月數
            ratio = round(up / total * 100, 1)
            out[f"{name}成長比率%"] = ratio
            out[f"{name}增長月數"] = f"{up}/{total}"     # 計數形式:增長月數/可比月數
            if w == CLASS_WINDOW:
                class_ratio, class_base = ratio, total
    out["_class_ratio"], out["_class_base"] = class_ratio, class_base

    last12 = rows[-12:]
    out["近12月樣態"] = "".join("✔" if r[2] > 0 else "✗" for r in last12)

    streak = 0                       # 連續成長月數(從最新往回)
    for r in reversed(rows):
        if r[2] > 0:
            streak += 1
        else:
            break
    out["連續成長月數"] = streak
    out[f"近{lookback//12}年平均年增%"] = round(float(np.mean([r[2] for r in recent])), 1)

    if want_detail:                  # 完整每月明細:(年月, 營收億, 年增%)
        out["_detail"] = [(f"{ym//100}/{ym%100:02d}", round(rv/1e8, 2), round(yoy, 1))
                          for ym, rv, yoy in recent]
    return out

def label(res):
    if not res or res.get("_class_base", 0) < RULES["min_base"]:
        return "— 資料不足/新上市"
    ratio = res.get("_class_ratio")
    if ratio is None:
        return "— 資料不足/新上市"
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
    print(f"總清單 {len(tickers)} 檔,待掃 {len(todo)} 檔,起始日 {start}")
    show_usage(api)
    print()

    sort_col = f"{dict(WINDOWS).get(CLASS_WINDOW,'近3年')}成長比率%"
    avg_col  = f"近{YEARS}年平均年增%"

    i = 0
    details = []                          # 明細模式用:累積 (代號, 年月, 營收億, 年增%)
    while i < len(todo):
        sid = todo[i]
        try:
            df = fetch_revenue(api, sid, start)
        except RateLimited:
            wait = seconds_to_next_hour()
            print(f"  ⏸ FinMind 每小時額度用罄 → 睡到下一個整點再續(約 {wait//60} 分 {wait%60} 秒)。"
                  f"不跳過、不標記 {sid},醒來重抓。")
            time.sleep(wait)
            continue                      # 同一檔重試,i 不前進、不寫 progress
        res = analyze(df, want_detail=WRITE_MONTHLY_DETAIL)
        row = {"代號": sid, "分類": label(res)}
        if res:
            if WRITE_MONTHLY_DETAIL and "_detail" in res:
                for ym, rv, yoy in res["_detail"]:
                    details.append({"代號": sid, "年月": ym, "營收(億)": rv, "年增%": yoy})
            row.update({k: v for k, v in res.items() if not k.startswith("_")})
        results.append(row)
        append_progress(row)              # 只有「真的抓到/真的查無資料」才記錄完成
        i += 1
        print(f"[{i}/{len(todo)}] {sid:6s} {row['分類']:14s} "
              f"{sort_col} {row.get(sort_col,'-')}% 連續 {row.get('連續成長月數','-')}月")
        time.sleep(RATE_SLEEP)

    # ---- 輸出 Excel(依分類窗的成長比率排序)----
    df = pd.DataFrame(results)
    if not df.empty:
        for c in [sort_col, "連續成長月數"] + [f"{n}成長比率%" for _, n in WINDOWS]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.sort_values([sort_col, "連續成長月數"], ascending=False, na_position="last")
        # 各窗並列「成長比率%」與「增長月數(增/可比,例 80/120)」
        win_cols = []
        for _, n in WINDOWS:                              # 近10年/5年/3年/1年
            win_cols += [f"{n}成長比率%", f"{n}增長月數"]
        cols = (["代號", "分類"]
                + win_cols
                + ["可比月數", "連續成長月數", "近12月樣態",
                   "最新月份", "最新月年增%", avg_col])
        df = df[[c for c in cols if c in df.columns]]

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="月營收年增掃描", index=False)
        if WRITE_MONTHLY_DETAIL and details:
            pd.DataFrame(details).to_excel(xw, sheet_name="每月營收年增明細", index=False)
    print(f"\n完成 → {OUTPUT}({len(df)} 檔)"
          + (f";另含 {len(details)} 列每月明細" if WRITE_MONTHLY_DETAIL and details else ""))
    if not df.empty and "分類" in df.columns:
        print(df["分類"].value_counts().to_string())


if __name__ == "__main__":
    main()
