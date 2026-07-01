# -*- coding: utf-8 -*-
"""
台股外資翻轉即時警報 tw_flip_alert.py
=======================================================================
每工作日晚間 20:30 台北跑 (FinMind 20:00 資料更新後):
  - 對 watchlist 每檔算今日 5d/20d 外資淨
  - 對比自己昨晚 snapshot, 偵測翻轉
  - **只在有實質翻轉時才寄 email** (減少噪音)

翻轉判定 (比 morning radar 嚴格):
  1. 🔴 5d 由正翻負且 > 200 萬股逆轉 (獲利了結訊號)
  2. 🟢 5d 由負翻正且 > 200 萬股逆轉 (底部確認)
  3. 📉 20d 單日巨變 (>= 800 萬股)
  4. 🚨 60d 20d 5d 三期同向反轉 (最強訊號)

Watchlist 只取 watchlist_tw.txt (核心 104 檔)

輸出:
  data/台股_翻轉警報.xlsx (只列 flip 檔)
  /tmp/flip_alert_body.html + subject.txt
  (若無 flip, 產空檔並 exit 0, workflow 不寄 email)
"""
import os, json, sys, time, requests
import pandas as pd
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
WATCHLIST_FILE = os.environ.get("WATCHLIST_FILE", "data/watchlist_tw.txt")
WORKERS = int(os.environ.get("WORKERS", "6"))
DST = "data/台股_翻轉警報.xlsx"
PREV_FILE = "data/flip_alert_prev.json"

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).strftime("%Y-%m-%d")
END = TODAY
START = (datetime.now(TPE) - timedelta(days=75)).strftime("%Y-%m-%d")

# ─── 翻轉判定 thresholds (千股) ───
FLIP_5D_REVERSE = 2000   # 5d 反向逆轉需超過 2000 千股規模
BIG_20D_CHANGE = 8000    # 20d 單日巨變需 8000 千股以上


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip() for t in env.replace(",", " ").split() if t.strip()]
        return list(dict.fromkeys([t for t in toks if t and not t.startswith("#")]))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip() for t in line.split() if t.strip())
        return list(dict.fromkeys(toks))
    return "2330 2454 2317".split()


def fm(dataset, data_id=None, start=None, end=None):
    p = {"dataset": dataset}
    if data_id: p["data_id"] = data_id
    if start: p["start_date"] = start
    if end: p["end_date"] = end
    if TOKEN: p["token"] = TOKEN
    for _ in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=20)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception: time.sleep(1)
    return pd.DataFrame()


def compute_radar(sid):
    df = fm("TaiwanStockInstitutionalInvestorsBuySellWide", data_id=sid, start=START, end=END)
    if df.empty: return sid, None
    df = df.sort_values("date")
    def s(sub, col): return int(sub[col].sum()) if col in sub.columns else 0
    def net_period(d):
        sub = df.tail(d) if len(df) >= d else df
        f = (s(sub,"Foreign_Investor_buy") + s(sub,"Foreign_Dealer_Self_buy")
             - s(sub,"Foreign_Investor_sell") - s(sub,"Foreign_Dealer_Self_sell"))
        t = s(sub,"Investment_Trust_buy") - s(sub,"Investment_Trust_sell")
        d_n = (s(sub,"Dealer_buy")+s(sub,"Dealer_self_buy")+s(sub,"Dealer_Hedging_buy")
               - s(sub,"Dealer_sell")-s(sub,"Dealer_self_sell")-s(sub,"Dealer_Hedging_sell"))
        return f, t, d_n
    f5, t5, d5 = net_period(5)
    f20, t20, d20 = net_period(20)
    f60, t60, d60 = net_period(60)
    return sid, {"外資5d":f5, "外資20d":f20, "外資60d":f60,
                 "投信5d":t5, "投信20d":t20,
                 "自營5d":d5, "自營20d":d20}


def detect_flips(cur_map, prev_map):
    """偵測翻轉訊號 (絕對狀態 + delta 二類, 絕對狀態不需要 prev)"""
    flips = []
    for sid, cur in cur_map.items():
        cur_f5 = cur.get("外資5d", 0) or 0
        cur_f20 = cur.get("外資20d", 0) or 0
        cur_f60 = cur.get("外資60d", 0) or 0

        events = []; severity = 0

        # ═══ 絕對狀態訊號 (首次跑也能觸發) ═══

        # 三期完全不同向 (5d/20d/60d 三個方向都不一樣) → 極大訊號
        signs = [1 if v > 0 else (-1 if v < 0 else 0) for v in (cur_f5, cur_f20, cur_f60)]
        max_abs = max(abs(cur_f5), abs(cur_f20), abs(cur_f60))
        if 1 in signs and -1 in signs and max_abs > 10_000_000:  # > 1 萬張
            events.append("🚨 三期方向不一致(5d↔20d↔60d)")
            severity += 3

        # 5d 與 20d 背離 (短期反轉中期趨勢)
        if cur_f5 * cur_f20 < 0 and max(abs(cur_f5), abs(cur_f20)) > 5_000_000:
            direction = "由買轉賣" if cur_f5 < 0 else "由賣轉買"
            events.append(f"🚨 5d↔20d 背離({direction})")
            severity += 2

        # 20d 與 60d 背離 (中期已在反轉長期趨勢)
        if cur_f20 * cur_f60 < 0 and max(abs(cur_f20), abs(cur_f60)) > 10_000_000:
            direction = "由買轉賣" if cur_f20 < 0 else "由賣轉買"
            events.append(f"🚨 20d↔60d 背離({direction})")
            severity += 2

        # ═══ Delta 訊號 (需要 prev 才能算) ═══
        prev = prev_map.get(sid)
        if prev:
            prv_f5 = prev.get("外資5d", 0) or 0
            prv_f20 = prev.get("外資20d", 0) or 0

            # 5d 由正翻負
            if prv_f5 > 0 and cur_f5 < 0 and abs(cur_f5 - prv_f5) > FLIP_5D_REVERSE * 1000:
                events.append("🔴 5d 翻紅"); severity += 2
            # 5d 由負翻正
            elif prv_f5 < 0 and cur_f5 > 0 and abs(cur_f5 - prv_f5) > FLIP_5D_REVERSE * 1000:
                events.append("🟢 5d 翻綠"); severity += 2

            # 20d 單日巨變
            d20_change = cur_f20 - prv_f20
            if abs(d20_change) >= BIG_20D_CHANGE * 1000:
                direction = "加碼" if d20_change > 0 else "減碼"
                emoji = "📈" if d20_change > 0 else "📉"
                events.append(f"{emoji} 20d 巨變({direction} {round(d20_change/1000,0):,.0f} 千)")
                severity += 1

        if events:
            flips.append({
                "代號": sid, "訊號": " / ".join(events), "嚴重度": severity,
                "外資5d(今)": cur_f5,
                "外資5d(昨)": (prev or {}).get("外資5d", None),
                "外資20d(今)": cur_f20,
                "外資20d(昨)": (prev or {}).get("外資20d", None),
                "外資60d(今)": cur_f60,
            })
    return sorted(flips, key=lambda x: -x["嚴重度"])


def build_email(flips, name_map, six_map=None):
    if not flips: return None, None
    n_severe = sum(1 for f in flips if f["嚴重度"] >= 3)
    n_normal = len(flips) - n_severe
    subject = f"🚨 [外資翻轉 {TODAY}] {len(flips)} 檔 (嚴重 {n_severe})"

    body = [f"""<html><body style='font-family:-apple-system,sans-serif;max-width:1100px'>
<h2>🚨 台股 外資翻轉即時警報 — {TODAY} 晚間</h2>
<p style='color:#666'>每工作日 20:30 台北跑, FinMind 20:00 更新資料後. 只有翻轉時才寄</p>
<p><b>{len(flips)}</b> 檔翻轉 | 🚨 嚴重 <b>{n_severe}</b> | 一般 {n_normal}</p>

<h3 style='margin-top:20px'>⚡ 翻轉清單 (按嚴重度)</h3>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%'>
<tr style='background:#f0f0f0'>
<th>嚴重</th><th>代號</th><th>名稱</th><th>六維</th>
<th>訊號</th>
<th>外資5d 今</th><th>昨</th>
<th>外資20d 今</th><th>昨</th>
<th>外資60d</th>
</tr>"""]

    for f in flips:
        code = str(f["代號"])
        severity_str = "🔴🔴🔴" if f["嚴重度"] >= 3 else ("🔴🔴" if f["嚴重度"] == 2 else "🔴")
        bg = "#ffcccc" if f["嚴重度"] >= 3 else ("#ffe5cc" if f["嚴重度"] == 2 else "")
        row_style = f"background:{bg}" if bg else ""
        six_info = six_map.get(code, "") if six_map else ""

        body.append(f"<tr style='{row_style}'>"
                    f"<td style='text-align:center'>{severity_str}</td>"
                    f"<td>{code}</td><td>{name_map.get(code,'')}</td>"
                    f"<td>{six_info}</td>"
                    f"<td><b>{f['訊號']}</b></td>"
                    f"<td style='text-align:right'>{round(f['外資5d(今)']/1000,0):,.0f}</td>"
                    f"<td style='text-align:right;color:#888'>{round(f['外資5d(昨)']/1000,0):,.0f}</td>"
                    f"<td style='text-align:right'>{round(f['外資20d(今)']/1000,0):,.0f}</td>"
                    f"<td style='text-align:right;color:#888'>{round(f['外資20d(昨)']/1000,0):,.0f}</td>"
                    f"<td style='text-align:right;color:#888'>{round(f['外資60d(今)']/1000,0):,.0f}</td>"
                    f"</tr>")

    body.append("</table>")

    body.append("""
<hr>
<h4>📌 訊號規則</h4>
<ul style='font-size:12px;color:#666'>
<li>🔴 5d 翻紅: 昨 5d > 0 今 5d < 0, 逆轉幅度 > 200 萬股</li>
<li>🟢 5d 翻綠: 昨 5d < 0 今 5d > 0, 逆轉幅度 > 200 萬股</li>
<li>📉/📈 20d 巨變: 單日 20d 累計變化 >= 800 萬股</li>
<li>🚨 三期背離: 5d 與 20d 方向相反 (短期反轉)</li>
<li>嚴重度: 🔴🔴🔴 = 3+ / 🔴🔴 = 2 / 🔴 = 1</li>
</ul>
<p style='color:#888;font-size:11px'>資料源: FinMind (20:00 更新) • 觀察 watchlist_tw.txt</p>
</body></html>""")

    return subject, "\n".join(body)


def load_prev():
    if not os.path.exists(PREV_FILE): return {}
    try:
        with open(PREV_FILE, encoding="utf-8") as f: return json.load(f)
    except: return {}


def save_prev(data):
    os.makedirs(os.path.dirname(PREV_FILE), exist_ok=True)
    with open(PREV_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def main():
    if not TOKEN: print("⚠️ 需 FINMIND_TOKEN"); sys.exit(1)

    codes = load_watchlist()
    print(f"外資翻轉警報 — {len(codes)} 檔 (平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(compute_radar, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            sid, d = fut.result()
            done += 1
            if d: results[str(sid)] = d
            if done % 20 == 0: print(f"  [{done}/{len(codes)}]")

    if not results:
        print("⚠️ 沒抓到資料"); sys.exit(1)

    # 名稱
    info_df = fm("TaiwanStockInfo")
    name_map = {}
    if not info_df.empty and "stock_id" in info_df.columns and "stock_name" in info_df.columns:
        info_df["stock_id"] = info_df["stock_id"].astype(str)
        name_map = dict(zip(info_df["stock_id"], info_df["stock_name"]))

    # 六維(若有)
    six_map = {}
    six_src = "data/台股100檔_六維交叉.xlsx"
    if os.path.exists(six_src):
        try:
            six = pd.read_excel(six_src)
            six["代號"] = six["代號"].astype(str)
            for _, r in six.iterrows():
                six_map[r["代號"]] = f"{r.get('評等','')} 品{r.get('品質','')} 分{r.get('六維分','')}"
        except: pass

    # 偵測翻轉
    prev = load_prev()
    prev_map = prev.get("map", {})
    flips = detect_flips(results, prev_map)
    print(f"翻轉檔數: {len(flips)}")

    # 存 snapshot
    save_prev({"date": TODAY, "map": results})

    if not flips:
        print("  無翻轉,不寄 email")
        # 產空 subject/body 讓 workflow 條件判斷
        with open("/tmp/flip_alert_subject.txt", "w", encoding="utf-8") as f: f.write("")
        with open("/tmp/flip_alert_body.html", "w", encoding="utf-8") as f: f.write("")
        return

    # 存 xlsx
    fdf = pd.DataFrame(flips)
    fdf["名稱"] = fdf["代號"].astype(str).map(name_map)
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    fdf.to_excel(DST, index=False)
    print(f"→ {DST}")

    # 組 email
    subject, body = build_email(flips, name_map, six_map)
    with open("/tmp/flip_alert_subject.txt", "w", encoding="utf-8") as f: f.write(subject)
    with open("/tmp/flip_alert_body.html", "w", encoding="utf-8") as f: f.write(body)
    print(f"→ {subject}")


if __name__ == "__main__":
    main()
