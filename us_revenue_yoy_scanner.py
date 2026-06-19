# -*- coding: utf-8 -*-
"""
美股季營收年增掃描器 (US Quarterly-Revenue YoY Scanner)
========================================================
為什麼是「季」不是「月」:美股公司只公布季度財報(無台灣那種法定月營收),
所以這支看「最近 3 年(12 季)每一季營收,是否較去年同期成長」。

資料來源:SEC EDGAR 官方 XBRL API(免費、免金鑰、歷史完整)。
  - companyfacts:一次抓一檔全部財報概念,從中取營收(us-gaap)。
  - 用財報原生的 fiscal year/period(fy/fp)+ 期間天數來組「單季」值:
      * 只收 ~3 個月(80-100 天)的期間 → 排除 10-Q 內的 6/9 個月 YTD 累計。
      * 用 fy/fp 標籤(非曆季 frame)→ 對非曆年公司(如 AAPL 9 月底結帳)也正確,
        不會被 SEC frame 的容差視窗丟季。
      * Q4 多數公司不單獨申報,改用「全年 − (Q1+Q2+Q3)」回填(同一 fy,口徑一致)。
        回填的季數會記在「Q4回填季數」欄,方便辨識哪些是合成值。
輸出:data/美股季營收年增掃描.xlsx

每檔算出:
  - 近12季成長比率% / 成長季數 / 可比基數
  - 連續成長季數(從最新季往回,連幾季 YoY 正)
  - 近4季成長、近4季樣態(✔/✗)
  - 最新季、最新季年增%、近12季平均年增%
  - 分類:🟢全期強勢 / 🔵多數成長 / 🟡中性 / 🔴轉弱衰退 / —資料不足

★★ 使用前務必做兩件事 ★★
  1. 把 USER_AGENT 改成「你的名字 你的email」——SEC 規定每個請求都要帶可辨識的
     User-Agent(含聯絡email),否則會被擋(403)。這是強制的。
  2. 準備清單:tickers_us.txt(一行一個美股代號,如 AAPL),或改下方 US_TICKERS。

清單來源(優先序):tickers_us.txt → tickers_us.csv → 內建 US_TICKERS。

速率:SEC 允許約 10 請求/秒,本程式預設 0.2 秒/檔(約 5/秒,禮貌值),
      撞到 403/429 會自動退避重試。掃幾百檔通常數分鐘內完成,沒有 FinMind 那種每小時額度。
"""

import os, sys, time, csv, re, json
import requests
import pandas as pd
import numpy as np
from datetime import datetime

# ---------- 設定 ----------
USER_AGENT = "ChangeMe yourname your_email@example.com"   # ★ 必改:SEC 要求帶聯絡資訊
OUTPUT     = "data/美股季營收年增掃描.xlsx"
PROGRESS   = "data/_us_revenue_scan_progress.csv"
CIK_CACHE  = "data/_sec_cik_map.json"

TICKER_FILE_TXT = "tickers_us.txt"
TICKER_FILE_CSV = "tickers_us.csv"

YEARS      = 3
LOOKBACK_Q = YEARS * 4            # 12 季
REQ_SLEEP  = 0.2                  # 每檔間隔秒(SEC 上限約 10/秒,取 5/秒禮貌值)
MAX_RETRY  = 4
RESUME     = True

# 內建清單(沒有 tickers_us.txt / .csv 時才用)
US_TICKERS = ["AAPL", "MSFT", "NVDA", "AMZN", "META", "GOOGL"]

# 嘗試的營收概念(依序合併;公司可能在不同年份換標籤,合併可補齊歷史)
REV_CONCEPTS = [
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "Revenues",
    "RevenueFromContractWithCustomerIncludingAssessedTax",
    "SalesRevenueNet",
]

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_URL       = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"

RULES = dict(strong_ratio=80, good_ratio=60, weak_ratio=40, min_base=6)

# 單季/全年的期間天數判斷(用來從原始財報期間挑出單季,排除 YTD 累計)
Q_MIN_DAYS, Q_MAX_DAYS = 80, 100      # 13 週=91 天,容 52/53 週制(可達 ~98 天)
Y_MIN_DAYS, Y_MAX_DAYS = 350, 380     # 52/53 週年度


def _headers():
    return {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"}


# ---------- 清單載入 ----------
def _parse_txt(path):
    out = []
    for line in open(path, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        for tok in line.replace(",", " ").replace("\t", " ").split():
            tok = re.sub(r"[^A-Z0-9.\-]", "", tok.strip().upper())  # 去掉黏進來的 } 等雜字元
            if tok:
                out.append(tok)
    return out

def _parse_csv(path):
    df = pd.read_csv(path, dtype=str)
    for col in df.columns:
        if str(col).strip().lower() in ("ticker", "symbol", "代號", "stock", "code"):
            return [str(x).strip().upper() for x in df[col].dropna()]
    return [str(x).strip().upper() for x in df.iloc[:, 0].dropna()]

def load_tickers():
    if os.path.exists(TICKER_FILE_TXT):
        t = _parse_txt(TICKER_FILE_TXT); print(f"從 {TICKER_FILE_TXT} 讀到 {len(t)} 檔"); return t
    if os.path.exists(TICKER_FILE_CSV):
        t = _parse_csv(TICKER_FILE_CSV); print(f"從 {TICKER_FILE_CSV} 讀到 {len(t)} 檔"); return t
    print(f"用內建清單 {len(US_TICKERS)} 檔"); return list(US_TICKERS)


# ---------- 代號 → CIK ----------
def load_cik_map():
    if os.path.exists(CIK_CACHE):
        try:
            return json.load(open(CIK_CACHE, encoding="utf-8"))
        except Exception:
            pass
    print("下載 SEC 代號→CIK 對照表...")
    r = requests.get(SEC_TICKERS_URL, headers=_headers(), timeout=30)
    r.raise_for_status()
    raw = r.json()
    m = {str(v["ticker"]).upper(): int(v["cik_str"]) for v in raw.values()}
    os.makedirs(os.path.dirname(CIK_CACHE), exist_ok=True)
    json.dump(m, open(CIK_CACHE, "w", encoding="utf-8"))
    print(f"  對照表 {len(m)} 檔,已快取")
    return m


# ---------- SEC 取數 ----------
def fetch_facts(cik):
    """抓 companyfacts;403/429 自動退避。回傳 dict 或 None。"""
    url = FACTS_URL.format(cik=cik)
    for attempt in range(MAX_RETRY):
        try:
            r = requests.get(url, headers=_headers(), timeout=30)
            if r.status_code == 404:
                return None                      # 該 CIK 無 XBRL 財報(如部分 ADR/ETF)
            if r.status_code in (403, 429):
                wait = 10 * (attempt + 1)
                print(f"  ! SEC 限制({r.status_code}),等 {wait}s 重試。"
                      f"(確認 USER_AGENT 已改成你的 email)")
                time.sleep(wait); continue
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"  ! CIK{cik} 取數失敗:{e}")
            time.sleep(2)
    return None

def _parse_d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()

def _concept_periods(facts):
    """
    從 companyfacts 抽出『單季(~3個月)』與『全年(~12個月)』營收。
    回傳 (quarters, annuals):
      quarters: {(fy, q): val}   q∈1..4,用財報自帶的 fiscal year/period(對非曆年公司也正確)
      annuals:  {fy: val}
    去重:同一(fy,period)可能出現在多次申報(原始/重編/比較欄),取 filed 最新者(重編優先);
         概念依 REV_CONCEPTS 優先序合併補洞(高優先序先到先得)。
    過濾:用期間天數判斷,只收單季(80-100天),排除 10-Q 內的 YTD 累計(6/9 個月)。
    """
    usg = (facts or {}).get("facts", {}).get("us-gaap", {})

    def collect(min_d, max_d, want_quarter):
        merged = {}
        for concept in REV_CONCEPTS:                      # 高優先序在前
            node = usg.get(concept)
            if not node:
                continue
            local = {}                                    # 本概念內:key -> (filed, val)
            for unit, items in node.get("units", {}).items():
                if unit != "USD":
                    continue
                for it in items:
                    s, e = it.get("start"), it.get("end")
                    val, fy = it.get("val"), it.get("fy")
                    fp, filed = it.get("fp"), it.get("filed", "")
                    if not (s and e) or val is None or fy is None:
                        continue
                    try:
                        days = (_parse_d(e) - _parse_d(s)).days
                    except Exception:
                        continue
                    if not (min_d <= days <= max_d):
                        continue
                    if want_quarter:
                        if fp not in ("Q1", "Q2", "Q3", "Q4"):
                            continue
                        key = (int(fy), int(fp[1]))
                    else:
                        key = int(fy)                     # 全年(fp 多為 FY,已用天數過濾)
                    prev = local.get(key)
                    if prev is None or filed > prev[0]:   # 取最新申報
                        local[key] = (filed, float(val))
            for key, (filed, val) in local.items():
                merged.setdefault(key, val)               # 概念優先序:先到先得補洞
        return merged

    quarters = collect(Q_MIN_DAYS, Q_MAX_DAYS, True)
    annuals  = collect(Y_MIN_DAYS, Y_MAX_DAYS, False)
    return quarters, annuals

def quarterly_revenue(facts):
    """
    取單季營收,並用『全年 − (Q1+Q2+Q3)』回填多數公司缺漏的 Q4。
    回傳 (rev, backfilled):
      rev:        {(fy, q): val}
      backfilled: set,內含被合成的 (fy, 4)
    """
    if not facts:
        return {}, set()
    quarters, annuals = _concept_periods(facts)
    backfilled = set()
    for fy in {y for (y, q) in quarters}:
        if (fy, 4) in quarters:                           # 已有實報 Q4,不動
            continue
        if all((fy, q) in quarters for q in (1, 2, 3)) and fy in annuals:
            sum3 = quarters[(fy, 1)] + quarters[(fy, 2)] + quarters[(fy, 3)]
            q4 = annuals[fy] - sum3
            avg3 = sum3 / 3
            # 合理性檢查:Q4 須為正、且落在三季平均的 0.2~5 倍(濾掉口徑不一致的異常合成值)
            if q4 > 0 and (avg3 <= 0 or 0.2 * avg3 <= q4 <= 5 * avg3):
                quarters[(fy, 4)] = q4
                backfilled.add((fy, 4))
    return quarters, backfilled


# ---------- 分析 ----------
def analyze(rev, backfilled=frozenset(), lookback=LOOKBACK_Q):
    if not rev:
        return None
    rows = []                                   # (fy, quarter, value, yoy%)
    for (y, q) in sorted(rev):
        prev = rev.get((y - 1, q))              # 同 fiscal 季的去年同期
        if prev and prev != 0:
            rows.append((y, q, rev[(y, q)], (rev[(y, q)] - prev) / abs(prev) * 100))
    if not rows:
        return None

    recent = rows[-lookback:]
    n_base = len(recent)
    n_grow = sum(1 for r in recent if r[3] > 0)
    n_bf   = sum(1 for r in recent if (r[0], r[1]) in backfilled)

    last4 = rows[-4:]
    g4 = sum(1 for r in last4 if r[3] > 0)
    pat4 = "".join("✔" if r[3] > 0 else "✗" for r in last4)

    streak = 0
    for r in reversed(rows):
        if r[3] > 0:
            streak += 1
        else:
            break

    latest = rows[-1]
    return {
        "最新季":           f"{latest[0]}Q{latest[1]}",
        "最新季年增%":      round(latest[3], 1),
        "近12季可比基數":    n_base,
        "近12季成長季數":    n_grow,
        "近12季成長比率%":  round(n_grow / n_base * 100, 1) if n_base else None,
        "近4季成長":        f"{g4}/{len(last4)}",
        "近4季樣態":        pat4,
        "連續成長季數":      streak,
        "近12季平均年增%":  round(float(np.mean([r[3] for r in recent])), 1),
        "Q4回填季數":       n_bf,
    }

def label(res):
    if not res or res["近12季可比基數"] < RULES["min_base"]:
        return "— 資料不足/新上市"
    ratio = res["近12季成長比率%"]
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

# 進度/輸出共用的固定欄位(各列 key 不一,需固定表頭避免欄位錯位)
PROG_FIELDS = ["代號", "分類", "最新季", "最新季年增%", "近12季可比基數", "近12季成長季數",
               "近12季成長比率%", "近4季成長", "近4季樣態", "連續成長季數",
               "近12季平均年增%", "Q4回填季數"]

def append_progress(row):
    os.makedirs(os.path.dirname(PROGRESS), exist_ok=True)
    new = not os.path.exists(PROGRESS)
    with open(PROGRESS, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=PROG_FIELDS, extrasaction="ignore", restval="")
        if new:
            w.writeheader()
        w.writerow(row)


# ---------- 主流程 ----------
def main():
    if "your_email@example.com" in USER_AGENT or USER_AGENT.startswith("ChangeMe"):
        print("⚠ 請先把程式最上方的 USER_AGENT 改成『你的名字 你的email』,"
              "否則 SEC 會擋(403)。改完再執行。")
        sys.exit(1)

    tickers = load_tickers()
    cikmap  = load_cik_map()
    done    = load_done()
    results = list(done.values())
    todo    = [t for t in tickers if t not in done]
    print(f"總清單 {len(tickers)} 檔,待掃 {len(todo)} 檔\n")

    for i, sym in enumerate(todo, 1):
        # SEC 對照表的 class 股用連字號(BRK-B),watchlist 常用點(BRK.B),兩種都試
        cik = cikmap.get(sym.upper()) or cikmap.get(sym.upper().replace(".", "-"))
        if cik is None:
            row = {"代號": sym, "分類": "— 查無CIK(非SEC財報/ETF/ADR)"}
        else:
            rev, backfilled = quarterly_revenue(fetch_facts(cik))
            res = analyze(rev, backfilled)
            row = {"代號": sym, "分類": label(res)}
            row.update(res or {})
        results.append(row)
        append_progress(row)
        print(f"[{i}/{len(todo)}] {sym:6s} {row['分類']:18s} "
              f"成長比率 {row.get('近12季成長比率%','-')}% 連續 {row.get('連續成長季數','-')}季")
        time.sleep(REQ_SLEEP)

    df = pd.DataFrame(results)
    if not df.empty:
        for c in ("近12季成長比率%", "連續成長季數", "近12季成長季數"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.sort_values(["近12季成長比率%", "連續成長季數", "近12季平均年增%"],
                            ascending=False, na_position="last")
        cols = ["代號", "分類", "近12季成長比率%", "近12季成長季數", "近12季可比基數",
                "連續成長季數", "近4季成長", "近4季樣態",
                "最新季", "最新季年增%", "近12季平均年增%", "Q4回填季數"]
        df = df[[c for c in cols if c in df.columns]]

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="美股季營收年增掃描", index=False)
    print(f"\n完成 → {OUTPUT}({len(df)} 檔)")
    if not df.empty and "分類" in df.columns:
        print(df["分類"].value_counts().to_string())


if __name__ == "__main__":
    main()
