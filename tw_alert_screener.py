# -*- coding: utf-8 -*-
"""
台股警戒掃描 tw_alert_screener.py
=======================================================================
一次性掃 6 類警戒清單, 跟 watchlist 交叉:

  1. 處置股票 (TaiwanStockDispositionSecuritiesPeriod, 近 180 天)
  2. 暫停交易 (TaiwanStockSuspended, 近 90 天)
  3. 暫停融券賣出 (TaiwanStockMarginShortSaleSuspension, 近 30 天)
  4. 借券爆量 (TaiwanStockSecuritiesLending 近 30 天累計, 對 watchlist 跑)
  5. 可轉債發行 (TaiwanStockConvertibleBondInfo, 全市場一次)
  6. 減資 (TaiwanStockCapitalReductionReferencePrice, 近 180 天)
  7. 下市 (TaiwanStockDelisting, 全市場一次)

Watchlist 來源: TICKERS env → data/watchlist_tw.txt → fallback

輸出 data/台股_警戒掃描.xlsx, 7+1 個分頁:
  - 警戒總覽 (watchlist 內有 alert 的, 整合)
  - 處置股票 / 暫停交易 / 暫停融券 / 借券爆量 / 可轉債發行 / 減資 / 下市
"""
import os, time, requests
import pandas as pd
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
DST = "data/台股_警戒掃描.xlsx"
WATCHLIST_FILE = "data/watchlist_tw.txt"
WORKERS = int(os.environ.get("WORKERS", "4"))
END = datetime.now().strftime("%Y-%m-%d")
D180 = (datetime.now() - timedelta(days=180)).strftime("%Y-%m-%d")
D90 = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
D30 = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")


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
    return "2330 2454 2317".split()


def fm(dataset, data_id=None, start=None, end=END):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for _ in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=30)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception:
            time.sleep(1)
    return pd.DataFrame()


def fetch_lending(sid):
    """個股 30 天借券累計量"""
    try:
        df = fm("TaiwanStockSecuritiesLending", data_id=sid, start=D30)
        if df.empty or "volume" not in df.columns:
            return sid, None
        total = int(df["volume"].sum())
        latest = int(df.sort_values("date").iloc[-1]["volume"])
        return sid, {"代號": sid, "30d 借券量": total, "最新日借券": latest}
    except Exception:
        return sid, None


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN")
    codes = load_watchlist()
    watch_set = set(str(c) for c in codes)
    print(f"台股警戒掃描 — {len(codes)} 檔, 抓全市場警戒清單後交叉")

    alerts_master = []  # 每筆 alert 一行, 後面 dedupe by 代號 + 類型

    # ─── 1. 處置股票 ───
    print("抓處置股票...")
    disp = fm("TaiwanStockDispositionSecuritiesPeriod", start=D180)
    disp_my = pd.DataFrame()
    if not disp.empty and "stock_id" in disp.columns:
        disp["stock_id"] = disp["stock_id"].astype(str)
        disp_my = disp[disp["stock_id"].isin(watch_set)].copy()
        for _, r in disp_my.iterrows():
            alerts_master.append({
                "代號": r["stock_id"], "名稱": r.get("stock_name", ""),
                "類型": "⚠️ 處置",
                "日期": r.get("date", ""),
                "說明": f"{r.get('condition','')} / {r.get('measure','')}",
                "起迄": f"{r.get('period_start','')} ~ {r.get('period_end','')}",
            })

    # ─── 2. 暫停交易 ───
    print("抓暫停交易...")
    susp = fm("TaiwanStockSuspended", start=D90)
    susp_my = pd.DataFrame()
    if not susp.empty and "stock_id" in susp.columns:
        susp["stock_id"] = susp["stock_id"].astype(str)
        susp_my = susp[susp["stock_id"].isin(watch_set)].copy()
        for _, r in susp_my.iterrows():
            alerts_master.append({
                "代號": r["stock_id"], "名稱": "",
                "類型": "🚫 暫停",
                "日期": r.get("date", ""),
                "說明": f"暫停{r.get('suspension_time','')}, 恢復{r.get('resumption_date','')}",
                "起迄": "",
            })

    # ─── 3. 暫停融券賣出 ───
    print("抓暫停融券...")
    ms = fm("TaiwanStockMarginShortSaleSuspension", start=D30)
    ms_my = pd.DataFrame()
    if not ms.empty and "stock_id" in ms.columns:
        ms["stock_id"] = ms["stock_id"].astype(str)
        ms_my = ms[ms["stock_id"].isin(watch_set)].copy()
        for _, r in ms_my.iterrows():
            alerts_master.append({
                "代號": r["stock_id"], "名稱": "",
                "類型": "📉 暫停融券",
                "日期": r.get("date", ""),
                "說明": r.get("reason", ""),
                "起迄": f"~ {r.get('end_date','')}",
            })

    # ─── 4. 借券爆量 (per stock) ───
    print(f"抓 watchlist 借券 (30d)...")
    lending_rows = []
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_lending, c): c for c in codes}
        for fut in as_completed(futs):
            sid, d = fut.result()
            if d: lending_rows.append(d)
    lend_df = pd.DataFrame(lending_rows)
    if not lend_df.empty:
        lend_df = lend_df.sort_values("30d 借券量", ascending=False)
        # 取 30d 借券 > 5 萬張的當警戒
        hot = lend_df[lend_df["30d 借券量"] > 50000]
        for _, r in hot.iterrows():
            alerts_master.append({
                "代號": r["代號"], "名稱": "",
                "類型": "🪦 借券爆量",
                "日期": END,
                "說明": f"30d {int(r['30d 借券量'])} 股, 最新日 {int(r['最新日借券'])} 股",
                "起迄": "",
            })

    # ─── 5. 可轉債發行 (近 180 天 issued) ───
    print("抓可轉債...")
    cb = fm("TaiwanStockConvertibleBondInfo")
    cb_my = pd.DataFrame()
    if not cb.empty:
        # cb_id 通常前面 4-5 碼是發行公司股號
        cb["__sid"] = cb["cb_id"].astype(str).str.extract(r"^(\d{4,5})")[0]
        cb_my = cb[cb["__sid"].isin(watch_set)].copy()
        # 取最近 180 天 InitialDateOfConversion
        if "InitialDateOfConversion" in cb_my.columns:
            cb_my = cb_my[cb_my["InitialDateOfConversion"] >= D180]
        for _, r in cb_my.iterrows():
            alerts_master.append({
                "代號": r["__sid"], "名稱": r.get("cb_name", ""),
                "類型": "🪙 可轉債(稀釋)",
                "日期": r.get("InitialDateOfConversion", ""),
                "說明": f"CB={r.get('cb_id','')} 金額={r.get('IssuanceAmount','')}",
                "起迄": f"~ {r.get('DueDateOfConversion','')}",
            })

    # ─── 6. 減資 ───
    print("抓減資...")
    cr_rows = []
    for sid in codes:
        cr = fm("TaiwanStockCapitalReductionReferencePrice", data_id=sid, start=D180)
        if cr.empty: continue
        for _, r in cr.iterrows():
            cr_rows.append({
                "代號": sid, "日期": r.get("date", ""),
                "原因": r.get("ReasonforCapitalReduction", ""),
                "減資前": r.get("ClosingPriceonTheLastTradingDay"),
                "減資後參考價": r.get("PostReductionReferencePrice"),
            })
            alerts_master.append({
                "代號": sid, "名稱": "",
                "類型": "✂️ 減資",
                "日期": r.get("date", ""),
                "說明": r.get("ReasonforCapitalReduction", ""),
                "起迄": "",
            })
    cr_df = pd.DataFrame(cr_rows)

    # ─── 7. 下市 ───
    print("抓下市...")
    de = fm("TaiwanStockDelisting")
    de_my = pd.DataFrame()
    if not de.empty and "stock_id" in de.columns:
        de["stock_id"] = de["stock_id"].astype(str)
        de_my = de[de["stock_id"].isin(watch_set)].copy()
        for _, r in de_my.iterrows():
            alerts_master.append({
                "代號": r["stock_id"], "名稱": r.get("stock_name", ""),
                "類型": "💀 下市",
                "日期": r.get("date", ""),
                "說明": "已下市", "起迄": "",
            })

    # ─── 警戒總覽 ───
    master_df = pd.DataFrame(alerts_master)

    # merge 體檢
    base = pd.DataFrame()
    for src in ["data/台股體檢總表.xlsx", "data/台股_體檢總表.xlsx"]:
        if not os.path.exists(src): continue
        try:
            h = pd.read_excel(src, sheet_name="體檢總表")
            h["代號"] = h["代號"].astype(str)
            base = h[[c for c in ["代號","名稱","產業","評等","品質總分"] if c in h.columns]]
            break
        except Exception: pass

    if not master_df.empty and not base.empty:
        master_df["代號"] = master_df["代號"].astype(str)
        master_df = master_df.merge(base[["代號","名稱","評等"]].rename(
            columns={"名稱":"名稱_base","評等":"評等"}), on="代號", how="left")
        master_df["名稱"] = master_df["名稱"].where(master_df["名稱"].astype(bool), master_df["名稱_base"])
        master_df = master_df.drop(columns=[c for c in ["名稱_base"] if c in master_df.columns])

    # 按代號排序, 同代號多種警戒會集中
    if not master_df.empty:
        master_df = master_df.sort_values(["代號","類型","日期"], ascending=[True, True, False])

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        if not master_df.empty:
            master_df.to_excel(xw, sheet_name="警戒總覽", index=False)
        else:
            pd.DataFrame([{"訊息": "watchlist 中沒有任何警戒, 全清白!"}]).to_excel(
                xw, sheet_name="警戒總覽", index=False)
        if not disp_my.empty: disp_my.to_excel(xw, sheet_name="處置股票", index=False)
        if not susp_my.empty: susp_my.to_excel(xw, sheet_name="暫停交易", index=False)
        if not ms_my.empty: ms_my.to_excel(xw, sheet_name="暫停融券", index=False)
        if not lend_df.empty: lend_df.to_excel(xw, sheet_name="借券排行", index=False)
        if not cb_my.empty: cb_my.to_excel(xw, sheet_name="可轉債發行", index=False)
        if not cr_df.empty: cr_df.to_excel(xw, sheet_name="減資", index=False)
        if not de_my.empty: de_my.to_excel(xw, sheet_name="下市", index=False)

    print(f"\n→ {DST}")
    if not master_df.empty:
        print(f"\n=== 警戒總覽 ({len(master_df)} 筆) ===")
        print(master_df.head(30).to_string(index=False))
        print(f"\n警戒類型分布:")
        print(master_df["類型"].value_counts().to_string())
    else:
        print("watchlist 中沒有任何警戒")


if __name__ == "__main__":
    main()
