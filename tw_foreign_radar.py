# -*- coding: utf-8 -*-
"""
台股持股外資動向雷達 tw_foreign_radar.py
=======================================================================
每工作日早上跑, 對 watchlist_tw.txt 每檔算:
  - 外資 / 投信 / 自營 5d / 20d / 60d 累計淨買賣 (千股)
  - 對比昨日 snapshot 找「翻轉」訊號
  - 綜合分 (依 20d 為主, 5d 加成)

輸出:
  data/台股_外資雷達.xlsx (總覽 + 翻轉警報)
  /tmp/foreign_radar_body.html + /tmp/foreign_radar_subject.txt

Watchlist: TICKERS env → data/watchlist_tw.txt → fallback
"""
import os, json, time, requests, sys
import pandas as pd
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
WATCHLIST_FILE = os.environ.get("WATCHLIST_FILE", "data/watchlist_tw.txt")
WORKERS = int(os.environ.get("WORKERS", "6"))
DST = "data/台股_外資雷達.xlsx"
PREV_FILE = "data/foreign_radar_prev.json"

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).strftime("%Y-%m-%d")
END = TODAY
START = (datetime.now(TPE) - timedelta(days=75)).strftime("%Y-%m-%d")  # 75 天保險


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


def fm(dataset, data_id, start=START, end=END):
    p = {"dataset": dataset, "data_id": data_id, "start_date": start, "end_date": end}
    if TOKEN: p["token"] = TOKEN
    for i in range(3):
        try:
            r = requests.get(BASE, params=p, timeout=20)
            if r.status_code == 429: time.sleep(3); continue
            if r.status_code != 200: return pd.DataFrame()
            return pd.DataFrame(r.json().get("data", []))
        except Exception: time.sleep(1)
    return pd.DataFrame()


def compute_radar(sid):
    """算某檔 5d/20d/60d 三大法人累計"""
    df = fm("TaiwanStockInstitutionalInvestorsBuySellWide", sid)
    if df.empty: return sid, None
    df = df.sort_values("date")
    def s(sub, col): return int(sub[col].sum()) if col in sub.columns else 0

    def net_period(d):
        sub = df.tail(d) if len(df) >= d else df
        f_net = s(sub, "Foreign_Investor_buy") - s(sub, "Foreign_Investor_sell")
        # 加 Foreign_Dealer_Self (外資自營, 2018 後才有)
        f_net += s(sub, "Foreign_Dealer_Self_buy") - s(sub, "Foreign_Dealer_Self_sell")
        t_net = s(sub, "Investment_Trust_buy") - s(sub, "Investment_Trust_sell")
        d_net = (s(sub, "Dealer_buy") + s(sub, "Dealer_self_buy") + s(sub, "Dealer_Hedging_buy")
                 - s(sub, "Dealer_sell") - s(sub, "Dealer_self_sell") - s(sub, "Dealer_Hedging_sell"))
        return f_net, t_net, d_net

    f5, t5, d5 = net_period(5)
    f20, t20, d20 = net_period(20)
    f60, t60, d60 = net_period(60)

    return sid, {
        "外資5d": f5, "投信5d": t5, "自營5d": d5,
        "外資20d": f20, "投信20d": t20, "自營20d": d20,
        "外資60d": f60, "投信60d": t60, "自營60d": d60,
        "資料天數": len(df),
    }


def foreign_signal(net_k):
    """外資 20d 判讀(千股)"""
    if net_k is None: return ("—", "")
    if net_k > 10000: return ("🚀 大買", "background:#a8e6a8")
    if net_k > 2000: return ("🟢 加碼", "background:#e5f5e5")
    if net_k > -2000: return ("🟡 中性", "background:#fffbe0")
    if net_k > -10000: return ("🟠 減碼", "background:#ffe5cc")
    return ("🔴 大賣", "background:#ffcccc")


def load_prev():
    if not os.path.exists(PREV_FILE): return {}
    try:
        with open(PREV_FILE, encoding="utf-8") as f: return json.load(f)
    except: return {}


def save_prev(data):
    os.makedirs(os.path.dirname(PREV_FILE), exist_ok=True)
    with open(PREV_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def detect_flips(cur_map, prev_map):
    """偵測翻轉訊號"""
    flips = []
    for sid, cur in cur_map.items():
        prev = prev_map.get(sid)
        if not prev: continue
        cur_f5 = cur.get("外資5d", 0) or 0
        prv_f5 = prev.get("外資5d", 0) or 0
        cur_f20 = cur.get("外資20d", 0) or 0
        prv_f20 = prev.get("外資20d", 0) or 0

        events = []
        # 5d 由正翻負
        if prv_f5 > 1000 and cur_f5 < -500:
            events.append("🔴 外資 5d 翻紅")
        # 5d 由負翻正
        elif prv_f5 < -1000 and cur_f5 > 500:
            events.append("🟢 外資 5d 翻綠")
        # 20d 巨變
        d20 = cur_f20 - prv_f20
        if abs(d20) >= 5000:
            direction = "加碼" if d20 > 0 else "減碼"
            events.append(f"{'📈' if d20>0 else '📉'} 20d 巨變({direction} {round(d20/1000,0):,.0f} 千股)")

        if events:
            flips.append({"代號": sid, "訊號": " / ".join(events),
                          "外資5d": cur_f5, "外資20d": cur_f20,
                          "外資5d(昨)": prv_f5, "外資20d(昨)": prv_f20})
    return flips


def build_email(df, flips, name_map):
    """組 HTML email"""
    # 訊號統計
    signals = df["訊號20d"].value_counts() if "訊號20d" in df.columns else {}
    n_buy_big = sum(1 for s in df.get("訊號20d", []) if isinstance(s, str) and "🚀" in s)
    n_buy = sum(1 for s in df.get("訊號20d", []) if isinstance(s, str) and "🟢" in s)
    n_sell = sum(1 for s in df.get("訊號20d", []) if isinstance(s, str) and "🟠" in s)
    n_sell_big = sum(1 for s in df.get("訊號20d", []) if isinstance(s, str) and "🔴" in s)

    subject = f"[外資雷達 {TODAY}] 🚀{n_buy_big} 🟢{n_buy} 🟠{n_sell} 🔴{n_sell_big}"
    if flips: subject += f" | 翻轉 {len(flips)}"

    body = [f"""<html><body style='font-family:-apple-system,sans-serif;max-width:1100px'>
<h2>🌍 台股 外資雷達 — {TODAY}</h2>
<p style='color:#666'>每工作日早上 8:30 更新。追蹤 watchlist 每檔 5/20/60d 三大法人累計買賣</p>
<p>🚀 <b>{n_buy_big}</b> 大買 | 🟢 <b>{n_buy}</b> 加碼 | 🟡 中性 | 🟠 <b>{n_sell}</b> 減碼 | 🔴 <b>{n_sell_big}</b> 大賣</p>"""]

    # ─── 翻轉警報 ───
    if flips:
        body.append("<div style='background:#fff8e0;padding:12px;border-left:4px solid #fa0'>")
        body.append(f"<h3 style='margin:0 0 8px 0'>⚡ 昨→今 翻轉警報 ({len(flips)} 檔)</h3>")
        body.append("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%'>")
        body.append("<tr style='background:#f0f0f0'><th>代號</th><th>名稱</th><th>訊號</th><th>外資5d</th><th>外資5d(昨)</th><th>外資20d</th></tr>")
        for f in flips[:20]:
            code = f["代號"]
            body.append(f"<tr><td>{code}</td><td>{name_map.get(str(code),'')}</td>"
                        f"<td><b>{f['訊號']}</b></td>"
                        f"<td style='text-align:right'>{round(f['外資5d']/1000,0):,.0f}</td>"
                        f"<td style='text-align:right;color:#888'>{round(f['外資5d(昨)']/1000,0):,.0f}</td>"
                        f"<td style='text-align:right'>{round(f['外資20d']/1000,0):,.0f}</td></tr>")
        body.append("</table></div>")

    # ─── 🚀 大買清單 ───
    big_buy = df[df["訊號20d"].str.contains("🚀", na=False)].sort_values("外資20d(千)", ascending=False)
    if not big_buy.empty:
        body.append("<h3 style='margin-top:24px'>🚀 大買清單 (20d 外資 > 1000 萬股)</h3>")
        body.append(_render_table(big_buy.head(15), name_map))

    # ─── 🔴 大賣清單 ───
    big_sell = df[df["訊號20d"].str.contains("🔴", na=False)].sort_values("外資20d(千)", ascending=True)
    if not big_sell.empty:
        body.append("<h3 style='margin-top:24px'>🔴 大賣清單 (20d 外資 < -1000 萬股)</h3>")
        body.append(_render_table(big_sell.head(15), name_map))

    # ─── 全表 TOP 20 (依 20d) ───
    body.append("<h3 style='margin-top:24px'>📊 全表 TOP 20 (外資 20d 累計)</h3>")
    top = df.sort_values("外資20d(千)", ascending=False).head(20)
    body.append(_render_table(top, name_map))

    body.append(f"""
<hr>
<p style='color:#888;font-size:11px'>
資料源: FinMind • 觀察 watchlist_tw.txt 每檔 5/20/60d 三大法人買賣<br>
訊號規則: 20d 外資 > 1000 萬股 🚀 / > 200 萬 🟢 / ± 200 萬 🟡 / < -200 萬 🟠 / < -1000 萬 🔴
</p></body></html>""")

    return subject, "\n".join(body)


def _render_table(df, name_map):
    h = ["<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px'>"]
    h.append("<tr style='background:#f0f0f0'>"
             "<th>代號</th><th>名稱</th>"
             "<th>外資5d</th><th>外資20d</th><th>外資60d</th>"
             "<th>投信20d</th><th>自營20d</th><th>訊號20d</th></tr>")
    for _, r in df.iterrows():
        code = r["代號"]
        style = r.get("__style", "")
        h.append(f"<tr>"
                 f"<td>{code}</td><td>{name_map.get(str(code),'')}</td>"
                 f"<td style='text-align:right'>{r.get('外資5d(千)','—'):,.0f}</td>"
                 f"<td style='text-align:right'>{r.get('外資20d(千)','—'):,.0f}</td>"
                 f"<td style='text-align:right;color:#888'>{r.get('外資60d(千)','—'):,.0f}</td>"
                 f"<td style='text-align:right'>{r.get('投信20d(千)','—'):,.0f}</td>"
                 f"<td style='text-align:right'>{r.get('自營20d(千)','—'):,.0f}</td>"
                 f"<td style='{style}'>{r.get('訊號20d','')}</td></tr>")
    h.append("</table>")
    return "\n".join(h)


def main():
    if not TOKEN: print("⚠️ 需 FINMIND_TOKEN"); sys.exit(1)

    codes = load_watchlist()
    print(f"外資雷達 — {len(codes)} 檔 (平行 {WORKERS})")

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
        print("⚠️ 沒抓到任何資料"); sys.exit(1)

    # 撈名稱
    print("撈 TaiwanStockInfo 名稱...")
    info_df = fm("TaiwanStockInfo", data_id="", start=None, end=None)
    name_map = {}
    if not info_df.empty and "stock_id" in info_df.columns and "stock_name" in info_df.columns:
        info_df["stock_id"] = info_df["stock_id"].astype(str)
        name_map = dict(zip(info_df["stock_id"], info_df["stock_name"]))

    # 組 DataFrame
    rows = []
    for sid, r in results.items():
        f20 = r.get("外資20d", 0) or 0
        sig, style = foreign_signal(round(f20/1000, 0))
        rows.append({
            "代號": sid, "名稱": name_map.get(str(sid), ""),
            "外資5d(千)": round((r.get("外資5d") or 0)/1000, 0),
            "外資20d(千)": round(f20/1000, 0),
            "外資60d(千)": round((r.get("外資60d") or 0)/1000, 0),
            "投信5d(千)": round((r.get("投信5d") or 0)/1000, 0),
            "投信20d(千)": round((r.get("投信20d") or 0)/1000, 0),
            "自營20d(千)": round((r.get("自營20d") or 0)/1000, 0),
            "訊號20d": sig,
            "__style": style,
        })
    df = pd.DataFrame(rows).sort_values("外資20d(千)", ascending=False)

    # 對比昨日, 偵測翻轉
    prev = load_prev()
    prev_map = prev.get("map", {})
    flips = detect_flips(results, prev_map)
    print(f"翻轉警報: {len(flips)} 檔")

    # 六維 merge (若有)
    six_dim_src = "data/台股100檔_六維交叉.xlsx"
    if os.path.exists(six_dim_src):
        try:
            six = pd.read_excel(six_dim_src)
            six["代號"] = six["代號"].astype(str)
            df = df.merge(six[["代號","評等","品質","分類","六維分"]], on="代號", how="left")
        except Exception: pass

    # 存 xlsx
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.drop(columns="__style", errors="ignore").to_excel(xw, sheet_name="總覽", index=False)
        if flips:
            fdf = pd.DataFrame(flips)
            fdf["名稱"] = fdf["代號"].astype(str).map(name_map)
            fdf.to_excel(xw, sheet_name="翻轉警報", index=False)

    print(f"→ {DST}")

    # 組 email
    subject, body = build_email(df, flips, name_map)
    with open("/tmp/foreign_radar_subject.txt", "w", encoding="utf-8") as f: f.write(subject)
    with open("/tmp/foreign_radar_body.html", "w", encoding="utf-8") as f: f.write(body)
    print(f"→ Subject: {subject}")

    # 存今日 snapshot (含日期)
    save_prev({"date": TODAY, "map": results})


if __name__ == "__main__":
    main()
