# -*- coding: utf-8 -*-
"""
台股籌碼儀表板 tw_chip_dashboard.py
=======================================================================
對 watchlist 跑 5+2 個 FinMind dataset, 整合成多分頁籌碼總表:

模組:
  1. 三大法人 → TaiwanStockInstitutionalInvestorsBuySellWide
     (外資 / 投信 / 自營 / 自營對沖 各別 90 天累計買賣)
  2. 外資持股 → TaiwanStockShareholding (持股率變化)
  3. 八大行庫 → TaiwanstockGovernmentBankBuySell (政府買盤)
  4. 散戶融資 → TaiwanStockMarginPurchaseShortSale (融資餘額 + 融券)
  5. 借券放空 → TaiwanStockSecuritiesLending (借券餘額)
  6. 處置警戒 → TaiwanStockDispositionSecuritiesPeriod (全市場一次抓)
  7. 暫停交易 → TaiwanStockSuspended (全市場一次抓)

Watchlist 來源: TICKERS env → data/watchlist_tw.txt → fallback

輸出 data/台股_籌碼儀表板.xlsx, 7 個分頁:
  - 總覽 (每檔籌碼分 + 訊號)
  - 三大法人
  - 外資持股
  - 八大行庫
  - 散戶融資
  - 借券放空
  - 警戒名單
"""
import os
import time
import requests
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股_籌碼儀表板.xlsx"
WATCHLIST_FILE = "data/watchlist_tw.txt"
WORKERS = int(os.environ.get("WORKERS", "4"))

# 抓近 N 天的籌碼資料
DAYS = int(os.environ.get("DAYS", "90"))
END = datetime.now().strftime("%Y-%m-%d")
START = (datetime.now() - timedelta(days=DAYS)).strftime("%Y-%m-%d")


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
        toks = [t for t in toks if t and not t.startswith("#")]
        if toks: return list(dict.fromkeys(toks))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip() for t in line.split() if t.strip())
        if toks: return list(dict.fromkeys(toks))
    return "2330 2454 2317 2308".split()


def fm_get(dataset, data_id=None, start_date=None, end_date=None):
    params = {"dataset": dataset}
    if data_id: params["data_id"] = data_id
    if start_date: params["start_date"] = start_date
    if end_date: params["end_date"] = end_date
    if TOKEN: params["token"] = TOKEN
    for attempt in range(3):
        try:
            r = requests.get(BASE, params=params, timeout=30)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 429: time.sleep(3 * (attempt+1)); continue
        if r.status_code != 200: return pd.DataFrame()
        try:
            j = r.json()
            return pd.DataFrame(j.get("data", []))
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def fetch_one(sid):
    """對一檔抓所有需要的 FinMind dataset"""
    try:
        out = {"代號": sid}

        # ─── 1. 三大法人(wide 表)90 天累計 ───
        inv = fm_get("TaiwanStockInstitutionalInvestorsBuySellWide",
                     data_id=sid, start_date=START, end_date=END)
        if not inv.empty:
            def sum_col(col):
                return int(inv[col].sum()) if col in inv.columns else 0
            # 外資
            f_buy = sum_col("Foreign_Investor_buy")
            f_sell = sum_col("Foreign_Investor_sell")
            f_net = f_buy - f_sell
            # 外資自營 (新類別)
            fd_buy = sum_col("Foreign_Dealer_Self_buy")
            fd_sell = sum_col("Foreign_Dealer_Self_sell")
            # 投信
            t_buy = sum_col("Investment_Trust_buy")
            t_sell = sum_col("Investment_Trust_sell")
            t_net = t_buy - t_sell
            # 自營商(自營 + 自營對沖 + 早期合併)
            d_buy = sum_col("Dealer_buy") + sum_col("Dealer_self_buy") + sum_col("Dealer_Hedging_buy")
            d_sell = sum_col("Dealer_sell") + sum_col("Dealer_self_sell") + sum_col("Dealer_Hedging_sell")
            d_net = d_buy - d_sell
            # 三大法人合計
            total_net = f_net + t_net + d_net
            out["外資90d淨"] = round(f_net / 1000, 0)  # 千股
            out["投信90d淨"] = round(t_net / 1000, 0)
            out["自營90d淨"] = round(d_net / 1000, 0)
            out["三大90d淨"] = round(total_net / 1000, 0)

        # ─── 2. 外資持股 ───
        sh = fm_get("TaiwanStockShareholding",
                    data_id=sid, start_date=START, end_date=END)
        if not sh.empty and "ForeignInvestmentSharesRatio" in sh.columns:
            sh = sh.sort_values("date")
            ratio_start = float(sh.iloc[0]["ForeignInvestmentSharesRatio"])
            ratio_end = float(sh.iloc[-1]["ForeignInvestmentSharesRatio"])
            out["外資持股%"] = round(ratio_end, 2)
            out["外資持股Δpp"] = round(ratio_end - ratio_start, 2)
            # 外資使用率(占可投資上限)
            if "ForeignInvestmentRemainRatio" in sh.columns:
                remain = float(sh.iloc[-1]["ForeignInvestmentRemainRatio"])
                out["外資餘額%"] = round(remain, 2)

        # ─── 3. 八大行庫(政府)90d ───
        gb = fm_get("TaiwanstockGovernmentBankBuySell",
                    data_id=sid, start_date=START, end_date=END)
        if not gb.empty:
            # 篩出該股
            gb_st = gb[gb["stock_id"] == sid] if "stock_id" in gb.columns else gb
            if not gb_st.empty:
                # 累計買 / 賣(股數 buy/sell)
                buy = int(gb_st["buy"].sum()) if "buy" in gb_st.columns else 0
                sell = int(gb_st["sell"].sum()) if "sell" in gb_st.columns else 0
                out["八大90d淨"] = round((buy - sell) / 1000, 0)

        # ─── 4. 散戶融資餘額 90d ───
        mg = fm_get("TaiwanStockMarginPurchaseShortSale",
                    data_id=sid, start_date=START, end_date=END)
        if not mg.empty:
            mg = mg.sort_values("date")
            if "MarginPurchaseTodayBalance" in mg.columns:
                m_start = float(mg.iloc[0]["MarginPurchaseTodayBalance"])
                m_end = float(mg.iloc[-1]["MarginPurchaseTodayBalance"])
                out["融資餘額(張)"] = round(m_end / 1000, 0)  # 千股≈張
                out["融資Δ%"] = round((m_end / m_start - 1) * 100, 1) if m_start > 0 else None
            if "ShortSaleTodayBalance" in mg.columns:
                s_end = float(mg.iloc[-1]["ShortSaleTodayBalance"])
                out["融券餘額(張)"] = round(s_end / 1000, 0)

        # ─── 5. 借券餘額 90d ───
        sl = fm_get("TaiwanStockSecuritiesLending",
                    data_id=sid, start_date=START, end_date=END)
        if not sl.empty and "volume" in sl.columns:
            sl = sl.sort_values("date")
            # 借券成交量(各日累計)
            vol_total = int(sl["volume"].sum())
            out["借券90d量"] = round(vol_total / 1000, 0)
            # 最新一日借券
            last_vol = int(sl.iloc[-1]["volume"])
            out["借券最新日"] = round(last_vol / 1000, 0)

        return sid, out
    except Exception as e:
        return sid, {"代號": sid, "__error": str(e)}


def chip_signal(row):
    """籌碼綜合分數 + 訊號 tag"""
    score = 0; tags = []
    # 外資 90d 淨買賣
    f = row.get("外資90d淨") or 0
    if f > 50000: score += 2; tags.append("🌍外資大買")
    elif f > 10000: score += 1; tags.append("🌍外資加碼")
    elif f < -50000: score -= 2; tags.append("🌍外資大賣")
    elif f < -10000: score -= 1; tags.append("🌍外資減碼")
    # 投信 90d
    t = row.get("投信90d淨") or 0
    if t > 10000: score += 1; tags.append("📈投信加碼")
    elif t < -10000: score -= 1; tags.append("📉投信減碼")
    # 八大行庫(政府)
    g = row.get("八大90d淨") or 0
    if g > 10000: score += 1; tags.append("🏦政府護盤")
    elif g < -10000: score -= 1; tags.append("🏦政府出清")
    # 外資持股率變化
    d = row.get("外資持股Δpp")
    if d is not None:
        if d > 2: score += 1; tags.append("📊外資佔比↑")
        elif d < -2: score -= 1; tags.append("📊外資佔比↓")
    # 散戶融資爆量(警戒)
    mg = row.get("融資Δ%")
    if mg is not None:
        if mg > 50: score -= 1; tags.append("⚠️散戶搶進(融資+50%)")
        elif mg < -30: tags.append("🧹散戶斷頭(融資-30%)")  # 中性, 可能築底
    # 借券大量(警戒)
    bl = row.get("借券90d量") or 0
    if bl > 100000: score -= 1; tags.append("🪦借券爆量(被放空)")
    return score, " ".join(tags) if tags else "—"


def main():
    if not TOKEN:
        print("⚠️ 未設 FINMIND_TOKEN(免費版速率慢, 建議付費版)")
    codes = load_watchlist()
    print(f"台股籌碼儀表板 — {len(codes)} 檔 × 5 個 dataset (平行 {WORKERS})")
    print(f"日期區間: {START} ~ {END} ({DAYS} 天)")

    # 個股 fetch
    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_one, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sid, data = fut.result()
            if data: results[sid] = data
            done += 1
            if done % 10 == 0: print(f"  [{done}/{len(codes)}]")

    # 全市場 fetch: 處置股 + 暫停交易
    print("抓全市場警戒名單...")
    disp_start = (datetime.now() - timedelta(days=180)).strftime("%Y-%m-%d")
    disp = fm_get("TaiwanStockDispositionSecuritiesPeriod", start_date=disp_start)
    susp_start = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    susp = fm_get("TaiwanStockSuspended", start_date=susp_start)

    # 體檢總表 merge
    base = pd.DataFrame()
    for src in ["data/台股體檢總表.xlsx", "data/台股_體檢總表.xlsx"]:
        if not os.path.exists(src): continue
        try:
            h = pd.read_excel(src, sheet_name="體檢總表")
            h["代號"] = h["代號"].astype(str)
            keep = [c for c in ["代號","名稱","產業","評等","品質總分"] if c in h.columns]
            base = h[keep]
            break
        except Exception: pass

    # 組總覽
    rows = []
    for sid in codes:
        r = results.get(sid, {})
        if not r: continue
        sc, tags = chip_signal(r)
        r2 = {k: v for k, v in r.items() if not k.startswith("__")}
        r2["籌碼分"] = sc
        r2["訊號"] = tags
        rows.append(r2)

    df = pd.DataFrame(rows)
    if not base.empty:
        df["代號"] = df["代號"].astype(str)
        df = df.merge(base, on="代號", how="left")
        front = [c for c in ["代號","名稱","產業","評等","品質總分",
                              "外資90d淨","投信90d淨","自營90d淨","三大90d淨",
                              "外資持股%","外資持股Δpp","外資餘額%",
                              "八大90d淨",
                              "融資餘額(張)","融資Δ%","融券餘額(張)",
                              "借券90d量","借券最新日",
                              "籌碼分","訊號"] if c in df.columns]
        rest = [c for c in df.columns if c not in front]
        df = df[front + rest]
    if "籌碼分" in df.columns:
        df = df.sort_values("籌碼分", ascending=False)

    # 各分頁 (sort 欄位缺失時不排序, 避免 Sponsor 層級資料缺漏)
    def safe_sort(d, col, **kw):
        return d.sort_values(col, **kw) if col in d.columns else d

    inv_sheet = df[[c for c in ["代號","名稱","評等","外資90d淨","投信90d淨","自營90d淨","三大90d淨","訊號"] if c in df.columns]].copy()
    inv_sheet = safe_sort(inv_sheet, "三大90d淨", ascending=False, na_position="last")

    sh_sheet = df[[c for c in ["代號","名稱","評等","外資持股%","外資持股Δpp","外資餘額%"] if c in df.columns]].copy()
    sh_sheet = safe_sort(sh_sheet, "外資持股Δpp", ascending=False, na_position="last")

    gov_sheet = df[[c for c in ["代號","名稱","評等","八大90d淨"] if c in df.columns]].copy()
    gov_sheet = safe_sort(gov_sheet, "八大90d淨", ascending=False, na_position="last")

    mar_sheet = df[[c for c in ["代號","名稱","評等","融資餘額(張)","融資Δ%","融券餘額(張)"] if c in df.columns]].copy()
    mar_sheet = safe_sort(mar_sheet, "融資Δ%", ascending=False, na_position="last")

    lend_sheet = df[[c for c in ["代號","名稱","評等","借券90d量","借券最新日"] if c in df.columns]].copy()
    lend_sheet = safe_sort(lend_sheet, "借券90d量", ascending=False, na_position="last")

    # 警戒名單(處置 + 暫停, 跟 watchlist 交叉)
    watch_set = set(str(c) for c in codes)
    alert_rows = []
    if not disp.empty and "stock_id" in disp.columns:
        disp["stock_id"] = disp["stock_id"].astype(str)
        disp_my = disp[disp["stock_id"].isin(watch_set)]
        for _, r in disp_my.iterrows():
            alert_rows.append({
                "代號": r["stock_id"], "名稱": r.get("stock_name", ""),
                "類型": "處置",
                "日期": r.get("date", ""),
                "原因": r.get("condition", ""),
                "措施": r.get("measure", ""),
                "起": r.get("period_start", ""), "迄": r.get("period_end", ""),
            })
    if not susp.empty and "stock_id" in susp.columns:
        susp["stock_id"] = susp["stock_id"].astype(str)
        susp_my = susp[susp["stock_id"].isin(watch_set)]
        for _, r in susp_my.iterrows():
            alert_rows.append({
                "代號": r["stock_id"], "名稱": r.get("stock_name", ""),
                "類型": "暫停",
                "日期": r.get("date", ""),
                "原因": r.get("note", ""),
                "措施": "", "起": "", "迄": "",
            })
    alert_sheet = pd.DataFrame(alert_rows)

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="總覽", index=False)
        inv_sheet.to_excel(xw, sheet_name="三大法人", index=False)
        sh_sheet.to_excel(xw, sheet_name="外資持股", index=False)
        gov_sheet.to_excel(xw, sheet_name="八大行庫", index=False)
        mar_sheet.to_excel(xw, sheet_name="散戶融資", index=False)
        lend_sheet.to_excel(xw, sheet_name="借券放空", index=False)
        if not alert_sheet.empty:
            alert_sheet.to_excel(xw, sheet_name="警戒名單", index=False)
        else:
            pd.DataFrame([{"訊息": "watchlist 中沒人在處置/暫停名單"}]).to_excel(
                xw, sheet_name="警戒名單", index=False)

    print(f"\n→ 已輸出 {DST}")
    print(f"分頁: 總覽 / 三大法人 / 外資持股 / 八大行庫 / 散戶融資 / 借券放空 / 警戒名單")
    print(f"\n=== 籌碼分 TOP 15 ===")
    show = [c for c in ["代號","名稱","評等","外資90d淨","投信90d淨","八大90d淨","外資持股Δpp","融資Δ%","籌碼分","訊號"] if c in df.columns]
    print(df[show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
