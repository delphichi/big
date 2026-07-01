# -*- coding: utf-8 -*-
"""
3:1 盈虧比入場計算工具 entry_calc.py
=======================================================================
基於 April1Stock 教學:
  Max Entry = (TP + 3 × SL) / 4
  盈虧比 = (TP - Entry) / (Entry - SL)

三種用法:

1. 單檔手動:
   python entry_calc.py NVDA 250 180
   → 給定 TP=250, SL=180, 算 Max Entry + 現價判讀

2. 單檔自動 (FMP):
   python entry_calc.py NVDA --auto
   → 用 分析師 targetLow 當 TP, MA200 或 52w低 當 SL

3. 批量掃 watchlist:
   python entry_calc.py --batch us     # 美股
   python entry_calc.py --batch tw     # 台股
   → 對每檔算 Max Entry, 列出可買清單

輸出:
  console 表格
  data/entry_signals_{market}.xlsx (batch 模式)
"""
import os, sys, time, json, requests
import pandas as pd
from datetime import datetime, timedelta

FMP_KEY = os.environ.get("FMP_API_KEY", "")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")
FMP_BASE = "https://financialmodelingprep.com/stable"
FM_BASE = "https://api.finmindtrade.com/api/v4/data"

RATIO = 3.0  # 3:1 盈虧比


def calc(tp, sl, price):
    """核心計算"""
    if tp <= sl:
        return {"錯誤": "TP 必須 > SL"}
    max_entry = (tp + RATIO * sl) / (RATIO + 1)
    upside = round((tp / price - 1) * 100, 1) if price else None
    downside = round((sl / price - 1) * 100, 1) if price else None
    actual_ratio = (tp - price) / (price - sl) if price and price > sl else None

    if price and price <= sl:
        verdict = "🚨 已破止損(SL 失守, 重新評估支撐)"
    elif price and price <= max_entry:
        verdict = "🟢 進場區(盈虧比 >= 3:1)"
    elif price and price <= max_entry * 1.05:
        verdict = "🟡 接近門檻(<5% 距離)"
    elif price and price <= max_entry * 1.15:
        verdict = "🟠 稍高(等回檔 5-15%)"
    else:
        verdict = "🔴 追高風險(掛限價單 Max Entry)"

    return {
        "TP 目標": round(tp, 2),
        "SL 止損": round(sl, 2),
        "Max Entry": round(max_entry, 2),
        "現價": round(price, 2) if price else None,
        "上檔 %": upside, "下檔 %": downside,
        "實際盈虧比": round(actual_ratio, 2) if actual_ratio else None,
        "判讀": verdict,
    }


def fmp_get(endpoint, **params):
    if not FMP_KEY: return None
    params["apikey"] = FMP_KEY
    for _ in range(3):
        try:
            r = requests.get(f"{FMP_BASE}/{endpoint}", params=params, timeout=15)
            if r.status_code == 429: time.sleep(2); continue
            if r.status_code != 200: return None
            return r.json()
        except: time.sleep(1)
    return None


def fm_get(dataset, data_id, start):
    p = {"dataset": dataset, "data_id": data_id, "start_date": start}
    if FINMIND_TOKEN: p["token"] = FINMIND_TOKEN
    for _ in range(3):
        try:
            r = requests.get(FM_BASE, params=p, timeout=20)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except: time.sleep(1)
    return pd.DataFrame()


def auto_us(sym):
    """美股: FMP 拉 SL/TP/price 自動算
    SL 用 max(MA200, 52w低 × 1.05)  → 保守取近期強支撐
    TP 用 targetLow (分析師最保守目標價) → 避免過樂觀
    """
    q = fmp_get("quote", symbol=sym) or [{}]
    q = q[0] if isinstance(q, list) and q else {}
    price = q.get("price")
    ma200 = q.get("priceAvg200")
    yr_low = q.get("yearLow")

    tgt = fmp_get("price-target-consensus", symbol=sym) or [{}]
    tgt = tgt[0] if isinstance(tgt, list) and tgt else {}
    tp = tgt.get("targetLow")

    if not (price and ma200 and yr_low and tp):
        return None

    # SL = 較高的支撐(避免用最低點,通常是恐慌價)
    sl = max(ma200, yr_low * 1.03)
    # 若 MA200 高於現價,改用 yr_low 加 5% 當保守 SL
    if ma200 >= price:
        sl = yr_low * 1.05

    return {"代號": sym, "SL": sl, "TP": tp, "現價": price,
            "MA200": ma200, "52w低": yr_low, "分析師TargetLow": tp}


def auto_tw(sid):
    """台股: FinMind 6M 高低 + PER 目標估
    SL = 近 6M 低點 × 1.03
    TP = 現價 × 1.5(較粗糙, 因 FinMind 沒分析師目標)
        更好: 用「歷史 PER 中位 × 未來 EPS」但需要更多資料
    """
    start = (datetime.now() - timedelta(days=180)).strftime("%Y-%m-%d")
    df = fm_get("TaiwanStockPriceAdj", sid, start)
    if df.empty or "close" not in df.columns: return None
    df = df.sort_values("date")
    price = float(df.iloc[-1]["close"])
    low = float(df["min"].min()) if "min" in df.columns else float(df["close"].min())
    high = float(df["max"].max()) if "max" in df.columns else float(df["close"].max())

    # PER 找歷史合理估值目標
    per = fm_get("TaiwanStockPER", sid, start)
    tp = high * 1.05  # 預設: 前高 + 5%
    if not per.empty and "PER" in per.columns:
        # 用歷史 PER 中位 * 未來 EPS 略微保守化
        pass  # 保持簡單, 用前高

    sl = low * 1.03  # 前低 + 3% 當支撐

    return {"代號": sid, "SL": sl, "TP": tp, "現價": price,
            "6M低": low, "6M高": high}


def batch_us():
    """對 US watchlist 全掃, 找進場區檔"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    wl = "data/watchlist_us.txt"
    codes = []
    if os.path.exists(wl):
        with open(wl, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if line: codes.extend(line.split())
    codes = list(dict.fromkeys(codes))
    print(f"批量掃 US {len(codes)} 檔...")

    results = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(auto_us, c): c for c in codes}
        for fut in as_completed(futs):
            d = fut.result()
            if d: results.append(d)

    rows = []
    for d in results:
        r = calc(d["TP"], d["SL"], d["現價"])
        rows.append({**d, **r})
    df = pd.DataFrame(rows).sort_values("實際盈虧比", ascending=False, na_position="last")

    # 存
    dst = "data/entry_signals_us.xlsx"
    os.makedirs("data", exist_ok=True)
    df.to_excel(dst, index=False)
    print(f"→ {dst}")

    # 顯示 進場區 + 接近門檻
    good = df[df["判讀"].str.contains("🟢|🟡", na=False)]
    broken = df[df["判讀"].str.contains("🚨", na=False)]
    print(f"\n=== 🟢🟡 進場區 / 接近門檻 ({len(good)} 檔) ===")
    cols = ["代號","現價","SL","TP","Max Entry","實際盈虧比","上檔 %","判讀"]
    print(good[cols].to_string(index=False))
    if len(broken):
        print(f"\n=== 🚨 已破止損 ({len(broken)} 檔) — 需重新評估支撐 ===")
        print(broken[cols].to_string(index=False))

    _write_email("us", good, cols, broken)


def batch_tw():
    """對 TW watchlist 全掃"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    wl = "data/watchlist_tw.txt"
    codes = []
    if os.path.exists(wl):
        with open(wl, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if line: codes.extend(line.split())
    codes = list(dict.fromkeys(codes))
    print(f"批量掃 TW {len(codes)} 檔...")

    results = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = {ex.submit(auto_tw, c): c for c in codes}
        for fut in as_completed(futs):
            d = fut.result()
            if d: results.append(d)

    rows = []
    for d in results:
        r = calc(d["TP"], d["SL"], d["現價"])
        rows.append({**d, **r})
    df = pd.DataFrame(rows).sort_values("實際盈虧比", ascending=False, na_position="last")

    dst = "data/entry_signals_tw.xlsx"
    os.makedirs("data", exist_ok=True)
    df.to_excel(dst, index=False)
    print(f"→ {dst}")

    good = df[df["判讀"].str.contains("🟢|🟡", na=False)]
    broken = df[df["判讀"].str.contains("🚨", na=False)]
    print(f"\n=== 🟢🟡 進場區 / 接近門檻 ({len(good)} 檔) ===")
    cols = ["代號","現價","SL","TP","Max Entry","實際盈虧比","判讀"]
    print(good[cols].to_string(index=False))
    if len(broken):
        print(f"\n=== 🚨 已破止損 ({len(broken)} 檔) — 需重新評估支撐 ===")
        print(broken[cols].to_string(index=False))

    _write_email("tw", good, cols, broken)


def _write_email(market, good, cols, broken=None):
    """產出 /tmp/entry_signals_{market}_subject.txt + _body.html
    有 🟢🟡 或 🚨 才寫 subject → workflow 用來判斷是否寄信
    """
    label = "美股" if market == "us" else "台股"
    broken_n = len(broken) if broken is not None else 0
    if len(good) == 0 and broken_n == 0:
        print(f"⚠️ {label} 無 🟢🟡🚨 檔, 不產 email")
        return

    subject_parts = []
    if len(good): subject_parts.append(f"{len(good)} 進場區/接近")
    if broken_n:  subject_parts.append(f"{broken_n} 已破損")
    subject = f"🎯 {label} 3:1 入場清單 — {' + '.join(subject_parts)}"
    with open(f"/tmp/entry_signals_{market}_subject.txt", "w", encoding="utf-8") as f:
        f.write(subject)

    green = good[good["判讀"].str.contains("🟢", na=False)]
    yellow = good[good["判讀"].str.contains("🟡", na=False)]
    sl_src = "MA200 or 52w低×1.03 (取較高)" if market == "us" else "近 6M 低點×1.03"
    tp_src = "分析師 targetLow (最保守目標)" if market == "us" else "近 6M 高×1.05 (較粗糙)"

    html = f"""<html><body style='font-family:-apple-system,sans-serif;max-width:900px'>
<h2>{label} 3:1 盈虧比入場清單</h2>
<p style='color:#666'>公式: Max Entry = (TP + 3×SL) / 4 → 現價 ≤ Max Entry 才符合 3:1</p>

<div style='background:#e8f7e8;padding:10px;border-left:4px solid #2a2'>
<h3>🟢 進場區 ({len(green)} 檔) — 現價已在 Max Entry 以下</h3>
{green[cols].to_html(index=False, escape=False) if len(green) else '<p>無</p>'}
</div>

<div style='background:#fffbe0;padding:10px;border-left:4px solid #fa0;margin-top:12px'>
<h3>🟡 接近門檻 ({len(yellow)} 檔) — 距 Max Entry &lt;5%, 可掛限價單</h3>
{yellow[cols].to_html(index=False, escape=False) if len(yellow) else '<p>無</p>'}
</div>

{f'''<div style="background:#fee;padding:10px;border-left:4px solid #c33;margin-top:12px">
<h3>🚨 已破止損 ({broken_n} 檔) — 現價已跌破 SL, 原設定失效</h3>
{broken[cols].to_html(index=False, escape=False)}
<p style="color:#c33;font-size:12px"><b>處理</b>: 不能再用原 SL/TP; 需重新找更低的支撐或直接排除</p>
</div>''' if broken_n else ''}

<div style='background:#f5f5ff;padding:12px;border-left:4px solid #55b;font-size:13px;margin-top:16px'>
<h4>📖 用法說明</h4>
<p><b>SL 來源</b>: {sl_src}</p>
<p><b>TP 來源</b>: {tp_src}</p>
<p><b>執行</b>: 🟢 直接掛限價單買; 🟡 掛 Max Entry 等回檔</p>
<p style='color:#c33'><b>注意</b>: SL/TP 自動估算, 實戰請自行核對技術支撐 + 基本面目標</p>
</div>
</body></html>"""
    with open(f"/tmp/entry_signals_{market}_body.html", "w", encoding="utf-8") as f:
        f.write(html)
    print(f"→ /tmp/entry_signals_{market}_subject.txt + _body.html")


def main():
    args = sys.argv[1:]

    # 批量
    if "--batch" in args:
        market = args[args.index("--batch") + 1] if len(args) > args.index("--batch") + 1 else "us"
        if market == "us": batch_us()
        elif market == "tw": batch_tw()
        else: print("市場: us / tw")
        return

    # 單檔手動: SYM TP SL
    if len(args) == 3:
        sym, tp, sl = args[0].upper(), float(args[1]), float(args[2])
        # 抓現價
        q = fmp_get("quote", symbol=sym) or [{}]
        price = (q[0] if isinstance(q, list) and q else {}).get("price")
        if not price:
            # 試台股
            start = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
            df = fm_get("TaiwanStockPriceAdj", sym, start)
            if not df.empty:
                price = float(df.sort_values("date").iloc[-1]["close"])
        print(f"\n=== {sym} ===")
        r = calc(tp, sl, price)
        for k, v in r.items(): print(f"  {k}: {v}")
        return

    # 單檔自動
    if len(args) >= 2 and args[1] in ("--auto", "-a"):
        sym = args[0].upper()
        d = auto_us(sym) or auto_tw(sym)
        if not d:
            print(f"⚠️ 抓不到 {sym} 資料"); return
        r = calc(d["TP"], d["SL"], d["現價"])
        print(f"\n=== {sym} (自動抓 SL/TP) ===")
        for k, v in {**d, **r}.items():
            if k not in ("代號",): print(f"  {k}: {v}")
        return

    print(__doc__)


if __name__ == "__main__":
    main()
