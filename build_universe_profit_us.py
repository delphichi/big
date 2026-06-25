#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股全市場獲利初篩(SEC EDGAR 版)  build_universe_profit_us.py
=====================================================================
台股用 FinMind 全掃;美股 FMP 免費僅 250/天(撐不起全市場),改用 SEC EDGAR
(免費、~10/秒)當淺層源 —— sec_facts.py 已能從 companyfacts 算 ROE/ROIC/含金量。

漏斗:SEC 全市場淺篩(免費)→ 標出遺珠 → 之後 FMP 只深抓通過的(省 250/天 額度)。

方法:
  cikmap = load_cik_map()      # 全美 ~1萬 tickers,1 次免費呼叫(已快取)
  逐檔 sec_facts(sym, cikmap)  # 每檔 1 次 SEC companyfacts
  篩:roe_pct ≥ 15 或 roic_pct ≥ 12(高資本效率);ocf_to_ni ≥ 0.8 才算真現金
  斷點續跑:data/_profit_cache_us/{ticker}.json;MAX_RUNTIME 到點退出
輸出:data/美股_獲利型候選.txt(代號/ROE/ROIC/含金量/FCF/是否在現有觀察名單)

★ 需 SEC_USER_AGENT(「名字 email」,SEC 強制,否則 403)。
注意:SEC XBRL 標記不一致 → 很多檔算不出(error/None),屬正常先天盲區。
"""
import os
import json
import time
from datetime import datetime

from us_revenue_yoy_scanner import load_cik_map           # 全美 CIK 對照(免費)
from sec_facts import sec_facts                            # 單檔 ROE/ROIC/含金量

OUT = "data/美股_獲利型候選.txt"
CACHE_DIR = "data/_profit_cache_us"
ROE_MIN = 15.0
ROIC_MIN = 12.0
CASH_MIN = 0.8                # ocf/ni ≥ 0.8 才算真現金(濾掉純帳面獲利)
MAX_RUNTIME_MIN = 50
REQ_SLEEP = 0.13             # SEC ≤10/秒,留緩衝


def _cache_path(t):
    return os.path.join(CACHE_DIR, f"{t}.json")


def load_cache(t):
    p = _cache_path(t)
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8"))
        except Exception:
            return None
    return None


def save_cache(t, obj):
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp = _cache_path(t) + ".tmp"
    json.dump(obj, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, default=float)
    os.replace(tmp, _cache_path(t))


def load_watch():
    """現有美股觀察名單(core + watch),用來標『遺珠』。"""
    have = set()
    for f in ("tickers_us_core.txt", "tickers_us.txt"):
        if os.path.exists(f):
            for line in open(f, encoding="utf-8"):
                s = line.split("#")[0].strip().upper()
                if s:
                    have.add(s)
    return have


def build_output(tickers):
    watch = load_watch()
    rows = []
    for t in tickers:
        c = load_cache(t)
        if not c or c.get("error") or c.get("skip"):
            continue
        # 排除權證/特別股/單位(ticker 含「-」或過長),只留乾淨普通股
        if "-" in t or len(t) > 5:
            continue
        roe = c.get("roe_pct"); roic = c.get("roic_pct")
        cash = c.get("ocf_to_ni"); gm = c.get("gross_margin"); fcf = c.get("fcf_ttm_b")
        # 任一比率爆表(>合理上限)= 微型股分母失真,整檔排除(不只單邊閘)
        if (roe is not None and roe > 80) or (roic is not None and roic > 60):
            continue
        roe_ok = roe is not None and roe >= ROE_MIN
        roic_ok = roic is not None and roic >= ROIC_MIN
        if not (roe_ok or roic_ok):
            continue
        if gm is None or gm <= 0:               # 虧本毛利不算好公司
            continue
        if cash is None or cash < CASH_MIN:     # 須有真現金(None 不再放行)
            continue
        rows.append({
            "代號": t, "ROE": roe, "ROIC": roic, "含金量": cash,
            "毛利": gm, "FCF(B)": c.get("fcf_ttm_b"),
            "遺珠": "" if t in watch else "✔",
            })
    if not rows:
        print("  尚無符合條件者(可能還在抓)"); return 0
    rows.sort(key=lambda r: (r["ROIC"] if r["ROIC"] is not None else -1), reverse=True)
    new = [r for r in rows if r["遺珠"]]
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(f"# 美股全市場獲利初篩(SEC)  符合 {len(rows)} 檔"
                f"(ROE≥{ROE_MIN} 或 ROIC≥{ROIC_MIN},且 含金量≥{CASH_MIN})\n")
        f.write(f"# 產生:{datetime.now().date().isoformat()}  遺珠(不在觀察名單)={len(new)} 檔\n")
        f.write("# 代號  ROE  ROIC  含金量  毛利  FCF(B)  遺珠\n")
        for r in rows:
            f.write(f"{r['代號']}  ROE={r['ROE']}  ROIC={r['ROIC']}  "
                    f"含金量={r['含金量']}  毛利={r['毛利']}  FCF={r['FCF(B)']}  {r['遺珠']}\n")
    print(f"  → {OUT}:符合 {len(rows)} 檔 / 遺珠 {len(new)} 檔")
    return len(rows)


def main():
    t0 = time.time()
    cikmap = load_cik_map()
    tickers = sorted(cikmap.keys())
    todo = [t for t in tickers if load_cache(t) is None]
    print(f"全美 {len(tickers)} 檔,已完成 {len(tickers)-len(todo)},待抓 {len(todo)}")
    build_output(tickers)                                   # 先用既有快取出一版

    done = 0
    for i, t in enumerate(todo, 1):
        if (time.time() - t0) / 60 > MAX_RUNTIME_MIN:
            print(f"達 {MAX_RUNTIME_MIN} 分上限,本輪 {done} 檔,存檔退出(下輪續抓)")
            break
        try:
            r = sec_facts(t, cikmap)
            if r.get("error"):
                save_cache(t, {"error": r["error"]})        # 標記非SEC財報/ETF/外國,免每輪重撞
            else:
                save_cache(t, {
                    "roe_pct": r.get("roe_pct"), "roic_pct": r.get("roic_pct"),
                    "ocf_to_ni": r.get("ocf_to_ni"), "gross_margin": r.get("gross_margin"),
                    "fcf_ttm_b": r.get("fcf_ttm_b"),
                })
            done += 1
            if i % 200 == 0:
                print(f"[{i}/{len(todo)}] 進度 {done} 檔 ...")
        except Exception as e:
            save_cache(t, {"error": str(e)[:50]})
        time.sleep(REQ_SLEEP)

    n = build_output(tickers)
    remain = len([t for t in tickers if load_cache(t) is None])
    print(f"完成本輪。符合 {n} 檔;尚待抓 {remain} 檔(下輪 CI 續跑)")


if __name__ == "__main__":
    main()
