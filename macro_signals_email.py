# -*- coding: utf-8 -*-
"""產 macro_signals email"""
import os, pandas as pd
from datetime import datetime, timezone, timedelta

SRC = "data/macro_signals.xlsx"
SUB_OUT = "/tmp/macro_subject.txt"
BODY_OUT = "/tmp/macro_body.html"
TODAY = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d")


def color(judge):
    s = str(judge)
    if "🔴" in s: return "background:#ffe5e5;color:#c00;font-weight:bold"
    if "🟠" in s: return "background:#fff4e0;color:#c60"
    if "🟡" in s: return "background:#fffbe0;color:#888"
    if "🟢" in s: return "background:#e5ffe5;color:#080"
    return ""


def main():
    if not os.path.exists(SRC):
        with open(SUB_OUT,"w") as f: f.write(f"[總經 {TODAY}] 信號表不存在")
        with open(BODY_OUT,"w") as f: f.write(f"<p>{SRC} 不存在</p>")
        return

    df = pd.read_excel(SRC)
    red = sum(1 for j in df.get("判讀", []) if isinstance(j,str) and "🔴" in j)
    yellow = sum(1 for j in df.get("判讀", []) if isinstance(j,str) and ("🟠" in j or "🟡" in j))
    green = sum(1 for j in df.get("判讀", []) if isinstance(j,str) and "🟢" in j)

    subject = f"[總經 {TODAY}] 🔴{red} 🟡{yellow} 🟢{green}"

    rows = ["<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:14px'>"]
    rows.append("<tr style='background:#f0f0f0'><th>信號</th><th>核心數據</th><th>判讀</th></tr>")
    for _, r in df.iterrows():
        name = r["信號"]
        judge = r.get("判讀", "")
        # 主數值組合
        keys = [k for k in r.index if k not in ("信號","判讀","狀態","資料源","誤差")]
        vals = " | ".join(f"{k}={r[k]}" for k in keys if pd.notna(r[k]))
        style = color(judge)
        rows.append(f"<tr><td><b>{name}</b></td><td>{vals}</td><td style='{style}'>{judge}</td></tr>")
    rows.append("</table>")
    table_html = "\n".join(rows)

    body = f"""
<html><body style='font-family:-apple-system,sans-serif'>
<h2>總經 15 信號燈 — {TODAY}</h2>
<p style='color:#888'>🔴 過熱/危險 {red} | 🟡 警戒 {yellow} | 🟢 健康 {green}</p>

{table_html}

<hr>
<h3>📌 判讀規則</h3>
<ul style='font-size:13px;color:#555'>
<li><b>房地產 Case-Shiller</b>:YoY > 10% 過熱 / 5-10% 健康 / < 0 修正</li>
<li><b>黃金</b>:> $2400 強(央行買盤)/ $2000-2400 區間 / < $2000 弱</li>
<li><b>原油 WTI</b>:> $90 過熱 / $60-85 區間 / < $60 弱</li>
<li><b>AUDUSD</b>:> 0.68 強(中國需求好) / 0.62-0.68 區間 / < 0.62 弱(中國降溫)</li>
<li><b>台灣出口</b>:YoY > 20% AI 強 / 5-20% 健康 / < 5 弱</li>
<li><b>SOX</b>:離 52 週高 < 3% 過熱 / -3 ~ -15% 健康 / < -15% 修正中</li>
<li><b>VIX</b>:> 25 恐慌 / 20-25 警戒 / < 20 平靜</li>
<li><b>CPI</b>:> 3% 升息壓力 / 2-3% 中性 / < 2% 降息空間</li>
<li><b>10Y 公債</b>:> 4.5% 緊縮股不利 / 3.5-4.5% 中性 / < 3.5% 寬鬆股利好</li>
<li><b>FedWatch</b>:年底隱含利率,越低降息預期越強</li>
<li><b>MOVE</b>:> 130 債市恐慌 / 100-130 警戒 / < 100 平靜</li>
<li><b>NY Fed 衰退機率</b>:> 30% 高警戒 / 15-30% 警戒 / < 15% 安全</li>
<li><b>OECD CLI</b>:>100 升=擴張 / <100 降=收縮(全球領先指標)</li>
<li><b>台灣 GDP YoY</b>:> 5% 強 / 2-5% 中性 / < 2% 弱</li>
<li><b>CSP 資本支出</b>:YoY > 50% AI 爆 / > 20% 強 / < 0 收縮</li>
</ul>

<hr>
<p style='color:#888;font-size:11px'>
資料源:FRED(美國)、EIA(原油庫存)、FMP(即時報價)三維交叉驗證 |
<a href='https://github.com/delphichi/big/blob/main/data/macro_signals.xlsx'>看完整 xlsx</a> |
<a href='https://github.com/delphichi/big/blob/main/data/macro_signals_log.csv'>歷史 CSV</a>
</p>
</body></html>
"""

    with open(SUB_OUT,"w",encoding="utf-8") as f: f.write(subject)
    with open(BODY_OUT,"w",encoding="utf-8") as f: f.write(body)
    print(f"→ Subject: {subject}")


if __name__ == "__main__":
    main()
