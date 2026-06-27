# -*- coding: utf-8 -*-
"""
從 台股PE監看表.xlsx 產生 email HTML 摘要 tw_pe_monitor_email.py
=======================================================================
輸出:
  /tmp/tw_email_subject.txt
  /tmp/tw_email_body.html
"""
import os
import pandas as pd
from datetime import datetime, timezone, timedelta

SRC = "data/台股PE監看表.xlsx"
SUB_OUT = "/tmp/tw_email_subject.txt"
BODY_OUT = "/tmp/tw_email_body.html"
TODAY = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d")


def fmt(v, dec=1):
    if pd.isna(v) or v == "": return "-"
    try: return f"{float(v):.{dec}f}"
    except: return str(v)


def color(a):
    if not isinstance(a, str): return ""
    if "過熱" in a: return "background:#ffe5e5;color:#c00"
    if "偏貴" in a: return "background:#fff4e0;color:#c60"
    if "合理" in a: return "background:#fffbe0;color:#888"
    if "便宜" in a or "成長" in a: return "background:#e5ffe5;color:#080"
    return ""


def table(df, cols, title):
    if df is None or len(df) == 0:
        return f"<h3>{title}</h3><p style='color:#888'>(無)</p>"
    out = [f"<h3>{title}({len(df)} 檔)</h3>",
           "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:13px'>",
           "<tr style='background:#f0f0f0'>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr>"]
    for _, r in df.iterrows():
        tds = []
        for c in cols:
            v = r.get(c)
            style = ""
            if c == "估值鬧鐘":
                style = color(v)
            elif c == "漲跌%":
                try:
                    f = float(v)
                    if f > 5: style = "color:#c00;font-weight:bold"
                    elif f < -5: style = "color:#080;font-weight:bold"
                except: pass
            cell = fmt(v) if isinstance(v, (int, float)) else (v if pd.notna(v) else "-")
            tds.append(f"<td style='{style}'>{cell}</td>")
        out.append("<tr>" + "".join(tds) + "</tr>")
    out.append("</table>")
    return "\n".join(out)


def main():
    if not os.path.exists(SRC):
        with open(SUB_OUT,"w") as f: f.write(f"[台股監看 {TODAY}] 監看表不存在")
        with open(BODY_OUT,"w") as f: f.write(f"<p>{SRC} 不存在</p>")
        return

    df = pd.read_excel(SRC, sheet_name="監看表")
    df["代號"] = df["代號"].astype(str)
    for c in ["品質總分","當前股價","PER即時","PBR即時","ForwardPE","EPS近3y%","PEG_使用","漲跌%"]:
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce")

    buy = df[df["估值鬧鐘"].isin(["🟢成長未反映","🟢未來便宜","🟢便宜"])].sort_values("品質總分", ascending=False)
    hot = df[df["估值鬧鐘"].isin(["🔴未來過熱","🔴過熱"])].sort_values("品質總分", ascending=False)
    exp = df[df["估值鬧鐘"].isin(["🟠未來偏貴","🟠偏貴"])].sort_values("品質總分", ascending=False)
    big = pd.DataFrame()
    if "漲跌%" in df.columns:
        big = df[df["漲跌%"].abs() >= 5].sort_values("漲跌%", ascending=False)

    cols = ["代號","名稱","評等","品質總分","當前股價","PER即時","ForwardPE","PEG_使用","估值鬧鐘"]
    if "漲跌%" in df.columns: cols.append("漲跌%")
    cols = [c for c in cols if c in df.columns]

    subject = f"[台股監看 {TODAY}] 🟢買{len(buy)} 🔴熱{len(hot)} ⚡異動{len(big)}"
    body = f"""
<html><body style='font-family:-apple-system,sans-serif'>
<h2>台股 PE / PEG / Forward PE 監看 — {TODAY}</h2>
<p style='color:#888'>監看 {len(df)} 檔 | 🟢買{len(buy)} | 🔴熱{len(hot)} | ⚡漲跌≥5% {len(big)}</p>

{table(big, cols, "⚡ 今日大幅異動 (漲跌≥5%)")}

{table(buy, cols, "🟢 買進信號(估值便宜 / 成長未反映)")}

{table(hot, cols, "🔴 過熱警示")}

{table(exp.head(15), cols, "🟠 未來偏貴 TOP 15")}

<hr>
<p style='color:#888;font-size:11px'>自動產生 |
<a href='https://github.com/delphichi/big/blob/main/data/%E5%8F%B0%E8%82%A1PE%E7%9B%A3%E7%9C%8B%E8%A1%A8.xlsx'>看完整 xlsx</a></p>
</body></html>
"""
    with open(SUB_OUT, "w", encoding="utf-8") as f: f.write(subject)
    with open(BODY_OUT, "w", encoding="utf-8") as f: f.write(body)
    print(f"→ Subject: {subject}")
    print(f"→ Body: {BODY_OUT}")


if __name__ == "__main__":
    main()
