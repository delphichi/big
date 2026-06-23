# -*- coding: utf-8 -*-
"""
持倉月營收追蹤 (Monthly Revenue Watchdog)
=======================================================================
讀 portfolio.yaml 中的持倉,用 FinMind 抓近 6 個月營收年增率,對照論點日誌的證偽條件:
  - 月營收 YoY 連續 N 個月為負  → 🚨 觸發證偽
  - 單月 YoY 嚴重轉負(≤門檻)   → ⚠️ 警示
  - 季毛利率跌破論點門檻         → 🚨 觸發證偽(讀 data/台股財報估值.xlsx『逐季毛利率』)
輸出 data/持倉月度追蹤.xlsx;任一檔有 🚨/⚠️ 觸發就寄 Gmail。
排程:每月 11 日 09:00 UTC (月營收 10 日多已公布)。
"""
import os, time, ssl, smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import yaml

TOKEN      = os.environ.get("FINMIND_TOKEN", "")
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "")
MAIL_TO    = os.environ.get("MAIL_TO", "") or GMAIL_USER
PORTFOLIO  = "portfolio.yaml"
FIN_FILE   = "data/台股財報估值.xlsx"
OUT        = "data/持倉月度追蹤.xlsx"


def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try: dl.login_by_token(api_token=TOKEN)
        except Exception as e: print("token失敗:", e)
    return dl


def get_recent_yoy(dl, sid, months=6):
    """抓近 months 個月的月營收年增%。需上一年同月有資料才能算。"""
    start = (datetime.now() - timedelta(days=550)).strftime("%Y-%m-%d")
    try:
        df = dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start)
    except Exception as e:
        print(f"  {sid} 抓月營收失敗:{e}"); return pd.DataFrame()
    if df is None or df.empty: return pd.DataFrame()
    df = df.sort_values("date").reset_index(drop=True)
    df["yoy"] = (df["revenue"].pct_change(12) * 100).round(1)
    out = df.dropna(subset=["yoy"]).tail(months)[["date", "revenue", "yoy"]]
    return out


def latest_q_gm(sid):
    """從財報估值 Excel 讀該檔最新季毛利率;失敗回 None。"""
    try:
        q = pd.read_excel(FIN_FILE, "逐季毛利率", index_col=0)
        for k in q.index:
            if str(k).split()[0] == sid:
                v = q.loc[k].dropna()
                if len(v): return round(float(v.iloc[-1]), 2)
    except Exception: pass
    return None


def evaluate(pos, yoys, qgm):
    """回 (狀態, 詳細, 是否觸發)。"""
    sid, nm = pos["sid"], pos["name"]
    flags, alerts = [], []
    if not yoys.empty:
        last = list(yoys["yoy"])
        # 連續轉負月數
        neg = 0
        for v in reversed(last):
            if v < 0: neg += 1
            else: break
        thr_n = pos.get("證偽_月營收", {}).get("連續轉負月數", 99)
        if neg >= thr_n:
            flags.append(f"🚨 月營收連續 {neg} 月 YoY<0(門檻 {thr_n})")
        # 單月嚴重轉負
        thr_s = pos.get("證偽_月營收", {}).get("單月嚴重轉負", -99)
        if last[-1] <= thr_s:
            alerts.append(f"⚠️ 最新月 YoY {last[-1]}% ≤ {thr_s}%")
    # 毛利率
    thr_gm = pos.get("證偽_毛利率", {}).get("跌破")
    if qgm is not None and thr_gm is not None and qgm < thr_gm:
        flags.append(f"🚨 季毛利率 {qgm}% 跌破 {thr_gm}%")
    if flags: status = "🚨 觸發證偽"
    elif alerts: status = "⚠️ 警示"
    else: status = "✓ 正常"
    detail = "; ".join(flags + alerts) if (flags or alerts) else "持續觀察"
    return status, detail, bool(flags)


def send_email(subject, html):
    if not (GMAIL_USER and GMAIL_PASS and MAIL_TO):
        print("未設 Gmail secrets,略過寄信"); return
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
    print(f"追蹤 {len(positions)} 檔持倉")
    dl = make_loader()
    rows, hits = [], []
    for pos in positions:
        sid = pos["sid"]
        yoys = get_recent_yoy(dl, sid, 6)
        qgm  = latest_q_gm(sid)
        status, detail, triggered = evaluate(pos, yoys, qgm)
        last_yoys = " / ".join(f"{v:+.0f}%" for v in yoys["yoy"].tolist()) if not yoys.empty else "—"
        rows.append({
            "代號": sid, "名稱": pos["name"], "分類": pos.get("分類", ""), "狀態": status,
            "近6月YoY": last_yoys, "最新季毛利": qgm,
            "毛利門檻": pos.get("證偽_毛利率", {}).get("跌破"),
            "詳細": detail, "註記": pos.get("額外註記", ""),
        })
        if triggered or status == "⚠️ 警示": hits.append((sid, pos["name"], status, detail))
        print(f"  {sid} {pos['name']:6s} {status} {detail}")
        time.sleep(0.4)
    df = pd.DataFrame(rows)
    os.makedirs("data", exist_ok=True)
    df.to_excel(OUT, index=False)
    today = datetime.now().strftime("%Y-%m-%d")
    print(f"\n→ {OUT}({len(df)}檔);觸發/警示 {len(hits)} 檔")
    if hits:
        html = f"<h3>{today} 持倉月度檢查</h3>" + df.to_html(index=False, border=1)
        send_email(f"【持倉月度檢查】{today} {len(hits)} 檔觸發/警示", html)


if __name__ == "__main__":
    main()
