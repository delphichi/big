# -*- coding: utf-8 -*-
"""
週報心動股 Email weekly_picks_email.py
=======================================================================
讀 4 個資料源, 組 HTML email:
  - data/美股95檔_六維交叉.xlsx (TOP10 + BOTTOM5)
  - data/台股100檔_六維交叉.xlsx (TOP10 + BOTTOM5)
  - data/macro_signals.xlsx (美國總經 15 信號)
  - data/台股_總經儀表板.xlsx (台股總經)

跟前週對比 (data/weekly_picks_prev.json) 找出:
  - 新進 TOP10 的股票
  - 跌出 TOP10 的股票
  - 六維分變化 ≥ 2 的

輸出:
  /tmp/weekly_subject.txt
  /tmp/weekly_body.html
"""
import os, json
import pandas as pd
from datetime import datetime, timezone, timedelta

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).strftime("%Y-%m-%d")
SUB_OUT = "/tmp/weekly_subject.txt"
BODY_OUT = "/tmp/weekly_body.html"
PREV_FILE = "data/weekly_picks_prev.json"

US_SRC = "data/美股95檔_六維交叉.xlsx"
TW_SRC = "data/台股100檔_六維交叉.xlsx"
MACRO_US = "data/macro_signals.xlsx"
MACRO_TW = "data/台股_總經儀表板.xlsx"


def load_six_dim(path):
    """讀六維交叉表, 回 sorted DataFrame"""
    if not os.path.exists(path): return pd.DataFrame()
    try:
        # 嘗試 "總覽" 分頁, 不存在則用第一個
        try:
            df = pd.read_excel(path, sheet_name="總覽")
        except Exception:
            df = pd.read_excel(path)  # 第一個 sheet
        df["代號"] = df["代號"].astype(str)
        if "六維分" in df.columns:
            return df.sort_values("六維分", ascending=False)
        return df
    except Exception as e:
        print(f"⚠️ {path}: {e}")
        return pd.DataFrame()


def load_prev():
    if not os.path.exists(PREV_FILE): return {}
    try:
        with open(PREV_FILE, encoding="utf-8") as f: return json.load(f)
    except Exception: return {}


def save_prev(data):
    os.makedirs(os.path.dirname(PREV_FILE), exist_ok=True)
    with open(PREV_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def diff_top10(cur_df, prev_dict, market_key):
    """比對前週: 新進 / 跌出"""
    if cur_df.empty: return [], [], []
    cur_top = cur_df.head(10)["代號"].astype(str).tolist()
    prev_top = prev_dict.get(market_key, {}).get("top10", [])
    new_in = [s for s in cur_top if s not in prev_top]
    fell_out = [s for s in prev_top if s not in cur_top]
    # 分數變化 (現有 - 上週)
    prev_scores = prev_dict.get(market_key, {}).get("scores", {})
    big_change = []
    for _, r in cur_df.iterrows():
        sym = str(r["代號"])
        old = prev_scores.get(sym)
        new = int(r["六維分"]) if pd.notna(r["六維分"]) else 0
        if old is not None:
            delta = new - int(old)
            if abs(delta) >= 2:
                big_change.append((sym, str(r.get("名稱","")), old, new, delta))
    return new_in, fell_out, big_change


def render_table(df, cols, title, color="#e5f5e5"):
    if df.empty: return ""
    h = [f"<h3 style='margin-top:24px'>{title}</h3>",
         f"<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;background:{color}'>"]
    h.append("<tr style='background:#f0f0f0'>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")
    for _, r in df.iterrows():
        row = "<tr>" + "".join(f"<td>{r.get(c,'')}</td>" for c in cols) + "</tr>"
        h.append(row)
    h.append("</table>")
    return "\n".join(h)


def render_macro(macro_us, macro_tw):
    html = ["<h3 style='margin-top:24px'>🌍 總經速覽</h3>"]
    if not macro_us.empty:
        html.append("<h4>美國 (15 信號)</h4>")
        # 數紅黃綠
        red = sum(1 for j in macro_us.get("判讀", []) if isinstance(j,str) and "🔴" in j)
        ylw = sum(1 for j in macro_us.get("判讀", []) if isinstance(j,str) and ("🟠" in j or "🟡" in j))
        grn = sum(1 for j in macro_us.get("判讀", []) if isinstance(j,str) and "🟢" in j)
        html.append(f"<p>🔴 {red} | 🟡 {ylw} | 🟢 {grn}</p>")
        html.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:12px'>")
        html.append("<tr style='background:#f0f0f0'><th>信號</th><th>判讀</th></tr>")
        for _, r in macro_us.iterrows():
            html.append(f"<tr><td>{r.get('信號','')}</td><td>{r.get('判讀','—')}</td></tr>")
        html.append("</table>")
    if not macro_tw.empty:
        html.append("<h4>台股 (景氣 + VIX + 匯率)</h4>")
        html.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:12px'>")
        html.append("<tr style='background:#f0f0f0'><th>信號</th><th>數值</th><th>判讀</th></tr>")
        for _, r in macro_tw.iterrows():
            html.append(f"<tr><td>{r.get('信號','')}</td><td>{r.get('數值','')}</td><td>{r.get('判讀','—')}</td></tr>")
        html.append("</table>")
    return "\n".join(html)


def main():
    us = load_six_dim(US_SRC)
    tw = load_six_dim(TW_SRC)
    macro_us = pd.read_excel(MACRO_US) if os.path.exists(MACRO_US) else pd.DataFrame()
    macro_tw = pd.read_excel(MACRO_TW, sheet_name="總覽") if os.path.exists(MACRO_TW) else pd.DataFrame()

    prev = load_prev()
    us_new, us_out, us_chg = diff_top10(us, prev, "us")
    tw_new, tw_out, tw_chg = diff_top10(tw, prev, "tw")

    # subject
    us_n = len(us); tw_n = len(tw)
    subject = f"[週報 {TODAY}] 心動股 美 {us_n} / 台 {tw_n}"
    if us_new or tw_new:
        subject += f" | 新進 美{len(us_new)} 台{len(tw_new)}"

    # body
    body = [f"""<html><body style='font-family:-apple-system,sans-serif;max-width:1100px'>
<h2>🎯 心動股週報 — {TODAY}</h2>
<p style='color:#666'>每週日自動更新。基於 6 維評分(成長 + 估值 + 體質 + 報酬 + 內部人 + 國會 / 籌碼 + 警戒)</p>"""]

    # 變動 highlight
    if us_new or tw_new or us_out or tw_out or us_chg or tw_chg:
        body.append("<div style='background:#fff8e0;padding:12px;border-left:4px solid #fa0'>")
        body.append("<h3 style='margin:0 0 8px 0'>📰 本週變動</h3>")
        if us_new: body.append(f"<p>🆕 <b>美股新進 TOP10</b>: {', '.join(us_new)}</p>")
        if us_out: body.append(f"<p>👋 <b>美股跌出 TOP10</b>: {', '.join(us_out)}</p>")
        if tw_new: body.append(f"<p>🆕 <b>台股新進 TOP10</b>: {', '.join(tw_new)}</p>")
        if tw_out: body.append(f"<p>👋 <b>台股跌出 TOP10</b>: {', '.join(tw_out)}</p>")
        if us_chg:
            body.append("<p><b>美股分數變動 ≥ 2:</b></p><ul>")
            for s,n,o,nw,d in us_chg[:10]:
                arr = "📈" if d > 0 else "📉"
                body.append(f"<li>{arr} {s} {n}: {o} → {nw} ({d:+d})</li>")
            body.append("</ul>")
        if tw_chg:
            body.append("<p><b>台股分數變動 ≥ 2:</b></p><ul>")
            for s,n,o,nw,d in tw_chg[:10]:
                arr = "📈" if d > 0 else "📉"
                body.append(f"<li>{arr} {s} {n}: {o} → {nw} ({d:+d})</li>")
            body.append("</ul>")
        body.append("</div>")

    # 美股 TOP/BOTTOM
    if not us.empty:
        body.append("<h2 style='margin-top:30px'>🇺🇸 美股</h2>")
        us_cols = [c for c in ["代號","名稱","評等","當前股價","DCF差%","1y超額%","3y超額%","六維分","六維訊號"] if c in us.columns]
        body.append(render_table(us.head(10), us_cols, "🟢 TOP 10", "#e5f5e5"))
        body.append(render_table(us.tail(5), us_cols, "🔴 BOTTOM 5", "#ffeaea"))

    # 台股 TOP/BOTTOM
    if not tw.empty:
        body.append("<h2 style='margin-top:30px'>🇹🇼 台股</h2>")
        tw_cols = [c for c in ["代號","名稱","評等","品質","分類","籌碼分","1y報酬%","1y超額%","六維分","六維訊號"] if c in tw.columns]
        body.append(render_table(tw.head(10), tw_cols, "🟢 TOP 10", "#e5f5e5"))
        body.append(render_table(tw.tail(5), tw_cols, "🔴 BOTTOM 5", "#ffeaea"))

    # 總經
    body.append(render_macro(macro_us, macro_tw))

    body.append(f"""
<hr style='margin-top:30px'>
<p style='color:#888;font-size:11px'>
資料源: FMP (美股) / FinMind (台股) / FRED (總經)<br>
<a href='https://github.com/delphichi/big/blob/main/data/美股95檔_六維交叉.xlsx'>美股六維交叉 xlsx</a> |
<a href='https://github.com/delphichi/big/blob/main/data/台股100檔_六維交叉.xlsx'>台股六維交叉 xlsx</a>
</p>
</body></html>""")

    with open(SUB_OUT, "w", encoding="utf-8") as f: f.write(subject)
    with open(BODY_OUT, "w", encoding="utf-8") as f: f.write("\n".join(body))

    # 存本週 snapshot 給下週比對
    snapshot = {"date": TODAY}
    if not us.empty:
        snapshot["us"] = {
            "top10": us.head(10)["代號"].astype(str).tolist(),
            "scores": {str(r["代號"]): int(r["六維分"]) for _, r in us.iterrows() if pd.notna(r.get("六維分"))},
        }
    if not tw.empty:
        snapshot["tw"] = {
            "top10": tw.head(10)["代號"].astype(str).tolist(),
            "scores": {str(r["代號"]): int(r["六維分"]) for _, r in tw.iterrows() if pd.notna(r.get("六維分"))},
        }
    save_prev(snapshot)

    print(f"→ Subject: {subject}")
    print(f"→ Body: {BODY_OUT}")


if __name__ == "__main__":
    main()
