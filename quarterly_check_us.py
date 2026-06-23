# -*- coding: utf-8 -*-
"""
美股持倉季度追蹤 (Quarterly Watchdog)
=======================================================================
讀 portfolio_us.yaml,用 yfinance 抓近 4 季季營收/季EPS YoY,對照證偽條件:
  - 單季營收 YoY ≤ 門檻 → 🚨
  - 單季 EPS  YoY ≤ 門檻 → 🚨
  - EPS 連續下滑 N 季    → 🚨
  - 季毛利率跌破論點門檻 → 🚨
輸出 data/美股_持倉季度追蹤.xlsx,觸發即寄 Gmail。
排程:每週一(財報旺季密集),避開過早抓不到新季資料。
"""
import os, time, ssl, smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import yaml

GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "")
MAIL_TO    = os.environ.get("MAIL_TO", "") or GMAIL_USER
PORTFOLIO  = "portfolio_us.yaml"
OUT        = "data/美股_持倉季度追蹤.xlsx"


def fetch_quarterly(sym):
    """回 dict:{rev:[..], eps:[..], gm:[..]}, 近 5 季由舊到新(供算 YoY 需近 5 個季點)。"""
    import yfinance as yf
    t = yf.Ticker(sym)
    qis = t.quarterly_income_stmt
    if qis is None or qis.empty:
        return None
    def take(*names):
        for n in names:
            if n in qis.index:
                s = qis.loc[n].dropna()
                s = s[sorted(s.index)]
                return [float(x) for x in s.values]
        return []
    rev = take("Total Revenue")
    gp  = take("Gross Profit")
    ni  = take("Net Income", "Net Income Common Stockholders")
    eps = take("Diluted EPS", "Basic EPS")
    gm  = [g/r*100 if r else None for g, r in zip(gp, rev)]
    return {"rev": rev, "eps": eps, "gm": gm}


def evaluate(pos, q):
    """回 (狀態, 詳細, 觸發?)"""
    flags = []
    if not q:
        return "❔抓不到", "yfinance 無季資料", False
    rev, eps, gm = q["rev"], q["eps"], q["gm"]
    # 算近 4 季 YoY
    rev_yoys = [(rev[i]/rev[i-4]-1)*100 if i >= 4 and rev[i-4] else None for i in range(len(rev))]
    eps_yoys = [(eps[i]/eps[i-4]-1)*100 if i >= 4 and eps[i-4] and eps[i-4] > 0 else None for i in range(len(eps))]
    rev_yoys = [r for r in rev_yoys if r is not None][-4:]
    eps_yoys = [e for e in eps_yoys if e is not None][-4:]

    # 營收
    thr_r = pos.get("證偽_營收", {}).get("單季YoY≤")
    if thr_r is not None and rev_yoys and rev_yoys[-1] <= thr_r:
        flags.append(f"🚨 季營收YoY {rev_yoys[-1]:.0f}% ≤ {thr_r}")
    # EPS
    thr_e = pos.get("證偽_EPS", {}).get("單季YoY≤")
    if thr_e is not None and eps_yoys and eps_yoys[-1] <= thr_e:
        flags.append(f"🚨 季EPS YoY {eps_yoys[-1]:.0f}% ≤ {thr_e}")
    thr_n = pos.get("證偽_EPS", {}).get("連續下滑季數")
    if thr_n and len(eps_yoys) >= thr_n:
        if all(eps_yoys[-i-1] < 0 for i in range(thr_n)):
            flags.append(f"🚨 EPS連續 {thr_n} 季下滑")
    # 毛利
    thr_g = pos.get("證偽_毛利率", {}).get("跌破")
    if thr_g and gm and gm[-1] is not None and gm[-1] < thr_g:
        flags.append(f"🚨 毛利率 {gm[-1]:.1f}% 跌破 {thr_g}")

    status = "🚨 觸發證偽" if flags else "✓ 正常"
    detail = "; ".join(flags) if flags else "持續觀察"
    return status, detail, bool(flags), rev_yoys, eps_yoys, (gm[-1] if gm else None)


def send_email(subject, html):
    if not (GMAIL_USER and GMAIL_PASS and MAIL_TO):
        print("未設 Gmail,略過"); return
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, MAIL_TO
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
            s.starttls(context=ssl.create_default_context())
            s.login(GMAIL_USER, GMAIL_PASS); s.send_message(msg)
        print("寄信→", MAIL_TO)
    except Exception as e: print("寄信失敗:", e)


def main():
    if not os.path.exists(PORTFOLIO):
        print(f"找不到 {PORTFOLIO}"); return
    cfg = yaml.safe_load(open(PORTFOLIO, encoding="utf-8"))
    positions = cfg.get("positions", [])
    if not positions:
        print("portfolio_us.yaml 無持倉,跳過"); return
    print(f"追蹤 {len(positions)} 檔美股持倉")
    rows, hits = [], []
    for pos in positions:
        sym = pos["sym"]
        q = fetch_quarterly(sym)
        st, dt, trig, ryoy, eyoy, gm = evaluate(pos, q)
        rows.append({
            "代號": sym, "名稱": pos["name"], "分類": pos.get("分類", ""), "狀態": st,
            "近4季營收YoY": " / ".join(f"{v:+.0f}%" for v in ryoy) if ryoy else "—",
            "近4季EPS YoY": " / ".join(f"{v:+.0f}%" for v in eyoy) if eyoy else "—",
            "最新季毛利": round(gm, 2) if gm is not None else None,
            "毛利門檻": pos.get("證偽_毛利率", {}).get("跌破"),
            "詳細": dt, "註記": pos.get("額外註記", ""),
        })
        if trig: hits.append((sym, pos["name"], dt))
        print(f"  {sym} {pos['name'][:16]:16s} {st} {dt}")
        time.sleep(0.5)
    df = pd.DataFrame(rows)
    os.makedirs("data", exist_ok=True)
    df.to_excel(OUT, index=False)
    today = datetime.now().strftime("%Y-%m-%d")
    print(f"\n→ {OUT}({len(df)}檔);觸發 {len(hits)} 檔")
    if hits:
        html = f"<h3>{today} 美股持倉季度檢查</h3>" + df.to_html(index=False, border=1)
        send_email(f"【美股持倉檢查】{today} {len(hits)} 檔觸發", html)


if __name__ == "__main__":
    main()
