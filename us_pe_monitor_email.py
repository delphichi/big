# -*- coding: utf-8 -*-
"""
從 美股PE監看表.xlsx 產生 email HTML 摘要 us_pe_monitor_email.py
=======================================================================
輸出:
  /tmp/email_subject.txt  - 信件標題(含當日重要訊號數)
  /tmp/email_body.html    - HTML body(買進/警示/漲跌異動分區)
"""
import os
import pandas as pd
from datetime import datetime, timezone, timedelta

SRC = "data/美股PE監看表.xlsx"
SUB_OUT = "/tmp/email_subject.txt"
BODY_OUT = "/tmp/email_body.html"

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).strftime("%Y-%m-%d")


def fmt(v, dec=1):
    if pd.isna(v) or v == "": return "-"
    try: return f"{float(v):.{dec}f}"
    except: return str(v)


def color_alarm(a):
    if not isinstance(a, str): return ""
    if "過熱" in a: return "background:#ffe5e5;color:#c00"
    if "偏貴" in a: return "background:#fff4e0;color:#c60"
    if "合理" in a: return "background:#fffbe0;color:#888"
    if "便宜" in a or "成長" in a: return "background:#e5ffe5;color:#080"
    return ""


def table(df, cols, title):
    if df is None or len(df) == 0:
        return f"<h3>{title}</h3><p style='color:#888'>(無)</p>"
    rows = []
    rows.append(f"<h3>{title}({len(df)} 檔)</h3>")
    rows.append("<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:13px'>")
    rows.append("<tr style='background:#f0f0f0'>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>")
    for _, r in df.iterrows():
        tds = []
        for c in cols:
            v = r.get(c)
            style = ""
            if c == "估值鬧鐘":
                style = color_alarm(v)
            elif c == "漲跌%":
                try:
                    f = float(v)
                    if f > 5: style = "color:#c00;font-weight:bold"
                    elif f < -5: style = "color:#080;font-weight:bold"
                except: pass
            tds.append(f"<td style='{style}'>{fmt(v) if isinstance(v,(int,float)) else (v if pd.notna(v) else '-')}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    rows.append("</table>")
    return "\n".join(rows)


def main():
    if not os.path.exists(SRC):
        with open(SUB_OUT, "w") as f: f.write(f"[美股監看 {TODAY}] 監看表不存在")
        with open(BODY_OUT, "w") as f: f.write(f"<p>{SRC} 不存在,跳過 email</p>")
        return

    df = pd.read_excel(SRC, sheet_name="監看表")
    for c in ["品質總分","當前股價","PER即時","ForwardPE即時","EPS3y%","PEG即時","漲跌%"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # 分類
    buy = df[df["估值鬧鐘"].isin(["🟢成長未反映","🟢未來便宜","🟢便宜"])].copy().sort_values("品質總分", ascending=False)
    hot = df[df["估值鬧鐘"].isin(["🔴未來過熱","🔴過熱"])].copy().sort_values("品質總分", ascending=False)
    expensive = df[df["估值鬧鐘"].isin(["🟠未來偏貴","🟠偏貴"])].copy().sort_values("品質總分", ascending=False)

    # 大漲大跌
    big_move = pd.DataFrame()
    if "漲跌%" in df.columns:
        big_move = df[df["漲跌%"].abs() >= 5].copy().sort_values("漲跌%", ascending=False)

    cols = ["代號","名稱","評等","品質總分","當前股價","PER即時","ForwardPE即時","PEG即時","估值鬧鐘"]
    if "漲跌%" in df.columns: cols.append("漲跌%")
    cols = [c for c in cols if c in df.columns]

    move_cnt = len(big_move)
    buy_cnt = len(buy)
    hot_cnt = len(hot)
    subject = f"[美股監看 {TODAY}] 🟢買{buy_cnt} 🔴熱{hot_cnt} ⚡異動{move_cnt}"

    body = f"""
<html><body style='font-family:-apple-system,sans-serif'>
<h2>美股 PE / PEG / Forward PE 監看 — {TODAY}</h2>
<p style='color:#888'>監看 {len(df)} 檔 | 🟢 買進信號 {buy_cnt} | 🔴 過熱警示 {hot_cnt} | ⚡ 漲跌≥5% {move_cnt}</p>

{table(big_move, cols, "⚡ 今日大幅異動 (漲跌≥5%)")}

{table(buy, cols, "🟢 買進信號(估值便宜 / 成長未反映)")}

{table(hot, cols, "🔴 過熱警示")}

{table(expensive.head(15), cols, "🟠 未來偏貴 TOP 15")}

<hr>
<p style='color:#888;font-size:11px'>自動產生 | <a href='https://github.com/delphichi/big/blob/main/data/%E7%BE%8E%E8%82%A1PE%E7%9B%A3%E7%9C%8B%E8%A1%A8.xlsx'>看完整 xlsx</a></p>
</body></html>
"""

    with open(SUB_OUT, "w", encoding="utf-8") as f: f.write(subject)
    with open(BODY_OUT, "w", encoding="utf-8") as f: f.write(body)
    print(f"→ Subject: {subject}")
    print(f"→ Body: {BODY_OUT} ({len(body)} chars)")


if __name__ == "__main__":
    main()
