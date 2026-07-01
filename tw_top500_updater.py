# -*- coding: utf-8 -*-
"""
台股市值前 500 大更新器 tw_top500_updater.py
=======================================================================
抓當前市值前 N 大股票, 更新 data/watchlist_tw_500.txt

抓法優先順序:
  1. TaiwanStockMarketValueWeight (最直接, 2024-10-30~)
  2. TaiwanStockMarketValue (per-stock, 全市場 bulk)
  3. Fallback: 用 TaiwanStockInfo 只取普通股(無市值排序)

用途:
  - 覆寫或另存 data/watchlist_tw_500.txt
  - 讓 tw_10y_financials.py 抓 500 檔 (而非 100)
  - 讓其他 tw_*_dashboard.py 對這 500 檔跑

跑法:
  FINMIND_TOKEN=xxx TOP_N=500 python tw_top500_updater.py
"""
import os, time, requests, sys
import pandas as pd
from datetime import datetime, timedelta

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
TOP_N = int(os.environ.get("TOP_N", "500"))
OUT_FILE = os.environ.get("OUT_FILE", "data/watchlist_tw_500.txt")


def fm(dataset, data_id=None, start=None, end=None):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for i in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=60)
            if r.status_code == 402:
                print(f"    ⚠️ 402 付費限制: {dataset}"); return pd.DataFrame()
            if r.status_code == 429:
                wait = 5 * (i+1); print(f"    429, wait {wait}s"); time.sleep(wait); continue
            if r.status_code != 200:
                print(f"    HTTP {r.status_code}: {dataset}"); return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception as e:
            print(f"    err: {e}"); time.sleep(2)
    return pd.DataFrame()


def latest_business_day(days_back=7):
    """回推最多 days_back 天找有資料的日期"""
    for i in range(days_back):
        d = (datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d")
        yield d


def source_market_value_weight():
    """TaiwanStockMarketValueWeight — 2024-10-30~
    真實欄位: rank, stock_id, stock_name, weight_per (float, %), date, type (twse/tpex)
    """
    print("1. 嘗試 TaiwanStockMarketValueWeight...")
    for d in latest_business_day():
        df = fm("TaiwanStockMarketValueWeight", start=d, end=d)
        if not df.empty:
            print(f"   ✓ 用日期 {d}, {len(df)} 檔, columns={list(df.columns)}")
            return df
    return pd.DataFrame()


def source_market_value():
    """TaiwanStockMarketValue bulk — 全市場市值"""
    print("2. 嘗試 TaiwanStockMarketValue bulk...")
    for d in latest_business_day():
        df = fm("TaiwanStockMarketValue", start=d, end=d)
        if not df.empty:
            print(f"   ✓ 用日期 {d}, {len(df)} 檔")
            return df
    return pd.DataFrame()


def source_price_times_shares():
    """Fallback: 用 TaiwanStockPrice bulk × 股數(來自 TaiwanStockShareholding)算市值"""
    print("3. Fallback: Price × Shares 自己算...")
    for d in latest_business_day():
        pr = fm("TaiwanStockPrice", start=d, end=d)
        if pr.empty: continue
        sh = fm("TaiwanStockShareholding", start=d, end=d)
        if sh.empty: continue
        pr["stock_id"] = pr["stock_id"].astype(str)
        sh["stock_id"] = sh["stock_id"].astype(str)
        merged = pr[["stock_id","close"]].merge(
            sh[["stock_id","NumberOfSharesIssued"]], on="stock_id", how="inner")
        merged["market_value"] = merged["close"] * merged["NumberOfSharesIssued"]
        merged = merged.dropna(subset=["market_value"])
        print(f"   ✓ 用日期 {d}, {len(merged)} 檔")
        return merged
    return pd.DataFrame()


def is_common_stock(sid):
    """普通股: 4 碼數字 且 非 00 開頭(ETF/期權)"""
    s = str(sid)
    if not s.isdigit(): return False
    if len(s) != 4: return False
    if s.startswith("00"): return False
    return True


def main():
    if not TOKEN:
        print("⚠️ 需 FINMIND_TOKEN"); sys.exit(1)

    print(f"=== 台股市值前 {TOP_N} 大更新 ({datetime.now().strftime('%Y-%m-%d')}) ===\n")

    # 抓資料 — 動態偵測排序欄
    def pick_col(df, candidates):
        for c in candidates:
            if c in df.columns: return c
        return None

    # 主源: TaiwanStockMarketValue (真實市值, 跨市場可比)
    df = source_market_value()
    weight_col = pick_col(df, ["market_value", "MarketValue"]) if not df.empty else None
    src_name = "MarketValue"
    # 備 1: MarketValueWeight (分市場 %, 跨市場不可比但 TWSE 前段仍準)
    if df.empty or weight_col is None:
        df = source_market_value_weight()
        weight_col = pick_col(df, ["weight_per", "weight"]) if not df.empty else None
        src_name = "MarketValueWeight (⚠️ 分市場權重)"
    # 備 2: Price × Shares 自算
    if df.empty or weight_col is None:
        df = source_price_times_shares()
        weight_col = "market_value" if not df.empty else None
        src_name = "Price×Shares 自算"
    if df.empty or weight_col is None:
        print("❌ 三源都失敗或找不到市值欄"); sys.exit(1)
    print(f"排序欄: {weight_col} (源: {src_name})")

    print(f"\n原始 {len(df)} 筆, 用 {weight_col} 排序")

    # 統一代號 str
    df["stock_id"] = df["stock_id"].astype(str)

    # 名稱: 優先用 dataset 自帶的 stock_name, 否則再抓 TaiwanStockInfo
    if "stock_name" in df.columns:
        df["名稱"] = df["stock_name"].fillna("")
    else:
        print("抓 TaiwanStockInfo 補名稱...")
        info = fm("TaiwanStockInfo")
        name_map = {}
        if not info.empty and "stock_name" in info.columns:
            info["stock_id"] = info["stock_id"].astype(str)
            name_map = dict(zip(info["stock_id"], info["stock_name"]))
        df["名稱"] = df["stock_id"].map(name_map).fillna("")
    df["is_common"] = df["stock_id"].apply(is_common_stock)
    common = df[df["is_common"]].copy()
    print(f"篩出普通股: {len(common)} 檔")

    top = common.sort_values(weight_col, ascending=False).head(TOP_N)
    print(f"\n市值/權重前 {TOP_N} 大 (前 10):")
    show = top[["stock_id","名稱",weight_col]].head(10)
    print(show.to_string(index=False))

    # 寫 watchlist
    os.makedirs(os.path.dirname(OUT_FILE) or ".", exist_ok=True)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(f"# 台股市值/權重前 {TOP_N} 大 (自動生成 {datetime.now().strftime('%Y-%m-%d')})\n")
        f.write(f"# 來源: FinMind, sort by {weight_col}\n")
        for _, r in top.iterrows():
            f.write(f"{r['stock_id']}  # {r['名稱']} ({weight_col}={r[weight_col]:.2f})\n")

    print(f"\n→ {OUT_FILE}")
    print(f"   {len(top)} 檔已寫入")

    # 順帶存 CSV 給 debug
    csv_out = OUT_FILE.replace(".txt", ".csv")
    top[["stock_id","名稱",weight_col]].to_csv(csv_out, index=False)
    print(f"→ {csv_out}")


if __name__ == "__main__":
    main()
