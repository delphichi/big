#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全市場獲利初篩(逐檔慢掃版)  build_universe_profit.py
=====================================================================
FinMind 不支援財報 date-bulk(無 data_id 回 400),故改逐檔抓 → 算 ROE/EPS新高,
補「產業白名單(build_universe 僅 144 檔)」的盲區(如帆宣:廠務工程不在能力圈關鍵字)。

每檔僅 2 次呼叫(損益 + 資產負債),~2000 檔 ≈ 4000 call,靠斷點續跑分多次 CI 累積:
  - 每檔算完即存 data/_profit_cache/{sid}.json(★ 快取需 commit,CI 容器才接得上)
  - MAX_RUNTIME_MIN 到點存檔退出,下次 CI 接著抓
  - 撞額度短重試→跳過,下輪補

篩選門檻(任一):
  ROE ≥ 15(近四季淨利 / 最新季權益)
  近四季EPS 創近12季高(獲利突破,轉機/成長未被產業標籤抓到)

輸出 data/獲利型候選.txt(代號/名稱/產業/ROE/EPS新高/是否在PICKS),標出「遺珠」。
★ 需 FINMIND_TOKEN。
"""
import os
import sys
import json
import time
from datetime import datetime
import pandas as pd
import numpy as np

TOKEN = os.environ.get("FINMIND_TOKEN", "")
START = "2022-01-01"                 # 取 ~3 年(算 TTM + EPS新高歷史)
OUT = "data/獲利型候選.txt"
CACHE_DIR = "data/_profit_cache"
ROE_MIN = 15.0
EPS_HIGH_LOOKBACK = 12               # 近四季EPS 是否為近 12 季最高
MAX_RUNTIME_MIN = 50                 # 單輪上限,到點存檔退出靠快取續跑
RATE_SLEEP = 0.3
RATE_WAIT_CAP = 90
MAX_RATE_RETRY = 2


def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try:
            dl.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("login 失敗(改匿名):", e)
    return dl


def _is_rate_limit(e):
    m = str(e).lower()
    return any(k in m for k in ("limit", "402", "429", "too many", "exceed", "request"))


def load_universe(dl):
    """全台股普通股(排除 ETF/權證/特別股/DR)。回傳 [(sid,name,industry)]。"""
    info = dl.taiwan_stock_info()
    skip_ind = {"ETF", "Index", "大盤", "ETN", "受益證券", "存託憑證", "創新板"}
    out = {}
    for _, r in info.iterrows():
        sid = str(r["stock_id"]); ind = str(r.get("industry_category", ""))
        # 普通股:4 碼數字、非 00 開頭(ETF)、產業非排除類
        if len(sid) != 4 or not sid.isdigit() or sid.startswith("00"):
            continue
        if ind in skip_ind:
            continue
        out[sid] = (r.get("stock_name", sid), ind)
    return [(s, n, i) for s, (n, i) in out.items()]


def _cache_path(sid):
    return os.path.join(CACHE_DIR, f"{sid}.json")


def load_cache(sid):
    p = _cache_path(sid)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def save_cache(sid, obj):
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp = _cache_path(sid) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, default=float)
    os.replace(tmp, _cache_path(sid))


def _series_by_type(df, types):
    """財報 long 格式 → {date: value}(取指定科目)。"""
    sub = df[df["type"].isin(types)].copy()
    if sub.empty:
        return {}
    sub["value"] = pd.to_numeric(sub["value"], errors="coerce")
    sub = sub.dropna(subset=["value"]).sort_values("date")
    return {str(d): float(v) for d, v in zip(sub["date"], sub["value"])}


def compute_one(dl, sid):
    """回傳 dict(ROE / 近四季EPS / EPS新高),或 None。每檔 2 call。"""
    inc = dl.taiwan_stock_financial_statement(stock_id=sid, start_date=START)
    bal = dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=START)
    if inc is None or inc.empty:
        return None

    eps = _series_by_type(inc, ["EPS"])
    ni = _series_by_type(inc, ["IncomeAfterTaxes", "ProfitAfterTax", "NetIncome"])
    eq = _series_by_type(bal, ["Equity", "TotalEquity", "EquityAttributableToOwnersOfParent"]) if bal is not None and not bal.empty else {}

    eps_q = [eps[d] for d in sorted(eps)]
    ni_q = [ni[d] for d in sorted(ni)]
    eq_latest = eq[sorted(eq)[-1]] if eq else None

    ttm_eps = float(np.nansum(eps_q[-4:])) if len(eps_q) >= 4 else (float(np.nansum(eps_q)) if eps_q else None)
    ttm_ni = float(np.nansum(ni_q[-4:])) if len(ni_q) >= 4 else None
    roe = (ttm_ni / eq_latest * 100) if (ttm_ni is not None and eq_latest and eq_latest > 0) else None

    # EPS新高:最新滾動四季 EPS 為近 N 季最高
    eps_high = False
    if len(eps_q) >= 8:
        ttms = [float(np.nansum(eps_q[i:i + 4])) for i in range(len(eps_q) - 3)]
        ttms = [x for x in ttms if pd.notna(x)]
        if ttms and ttms[-1] >= max(ttms[-EPS_HIGH_LOOKBACK:]):
            eps_high = True

    return {
        "ROE": round(roe, 1) if roe is not None else None,
        "EPS": round(ttm_eps, 2) if ttm_eps is not None else None,
        "EPS新高": eps_high,
    }


def build_output(universe):
    name = {s: n for s, n, i in universe}
    ind = {s: i for s, n, i in universe}
    try:
        from fetch_fundamentals_tw import PICKS
        picks = set(PICKS)
    except Exception:
        picks = set()

    rows = []
    for sid in name:
        c = load_cache(sid)
        if not c:
            continue
        roe, eps_high = c.get("ROE"), c.get("EPS新高")
        if (roe is not None and roe >= ROE_MIN) or eps_high:
            rows.append({
                "代號": sid, "名稱": name[sid], "產業": ind.get(sid, ""),
                "ROE": roe, "EPS": c.get("EPS"), "EPS新高": "✔" if eps_high else "",
                "已在PICKS": "✔" if sid in picks else "",
            })
    if not rows:
        print("尚無符合條件者(可能還在抓)"); return 0
    df = pd.DataFrame(rows).sort_values(["EPS新高", "ROE"], ascending=[False, False], na_position="last")
    new = df[df["已在PICKS"] == ""]

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(f"# 全市場獲利初篩  符合 {len(df)} 檔(ROE≥{ROE_MIN} 或 EPS創近{EPS_HIGH_LOOKBACK}季高)\n")
        f.write(f"# 產生:{datetime.now().date().isoformat()}  遺珠(不在PICKS)={len(new)} 檔\n")
        f.write("# 代號  名稱  產業  ROE  EPS  EPS新高  已在PICKS\n")
        for _, r in df.iterrows():
            f.write(f"{r['代號']}  {r['名稱']}  {r['產業']}  ROE={r['ROE']}  "
                    f"EPS={r['EPS']}  {r['EPS新高']}  {r['已在PICKS']}\n")
    print(f"  → {OUT}:符合 {len(df)} 檔 / 遺珠 {len(new)} 檔")
    return len(df)


def main():
    t0 = time.time()
    dl = make_loader()
    universe = load_universe(dl)
    todo = [s for s, n, i in universe if load_cache(s) is None]
    print(f"全市場普通股 {len(universe)} 檔,已完成 {len(universe)-len(todo)},待抓 {len(todo)}")
    build_output(universe)                       # 先用既有快取出一版

    done = 0
    for idx, sid in enumerate(todo, 1):
        if (time.time() - t0) / 60 > MAX_RUNTIME_MIN:
            print(f"達時間上限 {MAX_RUNTIME_MIN} 分,本輪抓 {done} 檔,存檔退出(下輪續抓)")
            break
        tries = 0
        while True:
            try:
                res = compute_one(dl, sid)
                save_cache(sid, res or {})
                done += 1
                if idx % 50 == 0:
                    print(f"[{idx}/{len(todo)}] {sid} 進度 {done} 檔 ...")
                break
            except Exception as e:
                if _is_rate_limit(e) and tries < MAX_RATE_RETRY:
                    tries += 1
                    print(f"  {sid} 撞額度,睡 {RATE_WAIT_CAP}s 重試 {tries}")
                    time.sleep(RATE_WAIT_CAP)
                    continue
                print(f"  {sid} 失敗跳過:{str(e)[:60]}")
                save_cache(sid, {})              # 存空,避免每輪重撞同一檔
                break
        time.sleep(RATE_SLEEP)

    n = build_output(universe)
    remain = len([s for s, n2, i in universe if load_cache(s) is None])
    print(f"完成本輪。全市場符合 {n} 檔;尚待抓 {remain} 檔(剩餘下輪 CI 續跑)")


if __name__ == "__main__":
    main()
