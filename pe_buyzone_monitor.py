# -*- coding: utf-8 -*-
"""
每日 PE 買入區間監看器 (Daily PE Buy-Zone Monitor)
==================================================
紀律:① 好公司(另由財報表篩) ② 好價格 ③ 每天監看 PE 是否落入買入區間。
本程式做第③步:對自訂清單,每天抓最新 PER,判斷「今日 PE 是否 ≤ 該股近5年 PE 的 P20」
(即 PE 位階 ≤ 20%,處於自身歷史最便宜的 1/5 區間)→ 落入就視為「買入區間」。

買入區間定義:PE 位階% ≤ BUY_PCTL(預設 20)。
資料來源:FinMind taiwan_stock_per_pbr(每檔 1 次呼叫,回傳近5年每日 PER/PBR/殖利率)。
清單來源:tickers_watch.txt(一行一個或逗號/空白分隔,# 為註解)。
輸出   :data/PE買入區間監看.xlsx(今日買入區間 + 全部監看);有訊號時寄 Gmail。

★ 寄信:需在 repo Secrets 設 GMAIL_USER(你的 gmail)、GMAIL_APP_PASSWORD(Google 應用程式密碼,
  非一般密碼)。可另設 MAIL_TO(預設寄給自己)。沒設則只產報告、不寄信。
"""

import os, time, smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import pandas as pd
import requests

# ---------- 設定 ----------
TOKEN       = os.environ.get("FINMIND_TOKEN", "")
WATCH_FILE  = "tickers_watch.txt"
OUTPUT      = "data/PE買入區間監看.xlsx"
BUY_PCTL    = 20                 # 買入區間:PE 位階 ≤ 此百分位
YEARS       = 5                  # PE 位階用近幾年分布
MIN_DAYS    = 120               # 至少要有的交易日數(資料太少不判斷)
REQ_SLEEP   = 0.3

GMAIL_USER  = os.environ.get("GMAIL_USER", "")
GMAIL_PASS  = os.environ.get("GMAIL_APP_PASSWORD", "")
MAIL_TO     = os.environ.get("MAIL_TO", "") or GMAIL_USER
EMAIL_ALWAYS = os.environ.get("EMAIL_ALWAYS", "") == "1"   # =1 則每天都寄(含「無訊號」摘要)


# ---------- 清單 / FinMind ----------
def load_watch():
    out = []
    if not os.path.exists(WATCH_FILE):
        return out
    for line in open(WATCH_FILE, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        for tok in line.replace(",", " ").replace("\t", " ").split():
            t = tok.strip()
            if t.isdigit():
                out.append(t)
    return list(dict.fromkeys(out))            # 去重保序

def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try:
            dl.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("token 登入失敗(改用免費額度):", e)
    return dl

def names_map(dl):
    try:
        info = dl.taiwan_stock_info()
        return {str(r["stock_id"]): str(r["stock_name"]) for _, r in info.iterrows()}
    except Exception:
        return {}

def get_per(dl, sid, start):
    try:
        return dl.taiwan_stock_per_pbr(stock_id=sid, start_date=start)
    except Exception:
        h = {"Authorization": f"Bearer {TOKEN}"} if TOKEN else {}
        r = requests.get("https://api.finmindtrade.com/api/v4/data",
                         params={"dataset": "TaiwanStockPER", "data_id": sid, "start_date": start},
                         headers=h, timeout=20)
        return pd.DataFrame(r.json().get("data", []))


# ---------- 分析:今日 PE 位階 / 買入門檻 ----------
def analyze(df):
    if df is None or df.empty or "PER" not in df.columns:
        return None
    df = df.sort_values("date")
    per = pd.to_numeric(df["PER"], errors="coerce")
    s = per[per > 0].dropna()                  # 濾掉虧損期的負/零 PER
    if len(s) < MIN_DAYS:
        return None
    cur  = float(s.iloc[-1])
    p20  = float(s.quantile(BUY_PCTL / 100))
    rank = round(float((s <= cur).mean() * 100))
    pbr  = pd.to_numeric(df.get("PBR"), errors="coerce").dropna()
    dy   = pd.to_numeric(df.get("dividend_yield"), errors="coerce").dropna()
    return {
        "PE現": round(cur, 2),
        "PE買入門檻(P20)": round(p20, 2),
        "PE位階%": rank,
        "距買入區間%": round((cur / p20 - 1) * 100, 1),     # >0:還要再跌這麼多%才進場;≤0:已在區間
        "PE5年低": round(float(s.min()), 2),
        "PE5年高": round(float(s.max()), 2),
        "PBR現": round(float(pbr.iloc[-1]), 2) if len(pbr) else None,
        "殖利率%": round(float(dy.iloc[-1]), 2) if len(dy) else None,
        "最新日": str(df["date"].iloc[-1]),
        "_buy": rank <= BUY_PCTL,
    }


# ---------- Email ----------
def send_email(subject, html):
    if not (GMAIL_USER and GMAIL_PASS and MAIL_TO):
        print("未設 GMAIL_USER/GMAIL_APP_PASSWORD,略過寄信")
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, MAIL_TO
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
            s.starttls(context=ssl.create_default_context())
            s.login(GMAIL_USER, GMAIL_PASS)
            s.send_message(msg)
        print("已寄出 email →", MAIL_TO)
    except Exception as e:
        print("寄信失敗:", e)


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    nm = names_map(dl)
    start = f"{datetime.now().year - YEARS}-01-01"
    watch = load_watch()
    print(f"監看 {len(watch)} 檔,買入區間 = PE 位階 ≤ {BUY_PCTL}%")
    rows = []
    for i, sid in enumerate(watch, 1):
        try:
            r = analyze(get_per(dl, sid, start))
        except Exception as e:
            print(f"  ! {sid} 失敗:{e}"); r = None
        if r:
            rows.append({"代號": sid, "名稱": nm.get(sid, sid), **r})
            tag = "★買入區間" if r["_buy"] else f"距{r['距買入區間%']}%"
            print(f"[{i}/{len(watch)}] {sid} {nm.get(sid,sid):6s} PE {r['PE現']:>7} 位階{r['PE位階%']:>3} {tag}")
        time.sleep(REQ_SLEEP)

    if not rows:
        print("無有效資料"); return
    df = pd.DataFrame(rows).sort_values("PE位階%")
    buy = df[df["_buy"]].drop(columns="_buy")
    allv = df.drop(columns="_buy")

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        buy.to_excel(xw, sheet_name="今日買入區間", index=False)
        allv.to_excel(xw, sheet_name="全部監看", index=False)

    today = datetime.now().strftime("%Y-%m-%d")
    if len(buy):
        cols = ["代號", "名稱", "PE現", "PE買入門檻(P20)", "PE位階%", "PE5年低", "殖利率%"]
        html = (f"<h3>{today} 進入買入區間(PE位階 ≤ {BUY_PCTL}%)共 {len(buy)} 檔</h3>"
                + buy[cols].to_html(index=False, border=1))
        send_email(f"【PE買入區間】{today} {len(buy)} 檔進場訊號", html)
    elif EMAIL_ALWAYS:
        near = allv.sort_values("距買入區間%").head(8)
        cols = ["代號", "名稱", "PE現", "PE買入門檻(P20)", "距買入區間%", "殖利率%"]
        html = (f"<h3>{today} 無股票進入買入區間</h3><p>最接近的(還需再跌%):</p>"
                + near[cols].to_html(index=False, border=1))
        send_email(f"【PE買入區間】{today} 今日無進場訊號", html)

    print(f"\n完成 → {OUTPUT};今日買入區間 {len(buy)} 檔 / 監看 {len(allv)} 檔")
    if len(buy):
        print(buy[["代號", "名稱", "PE現", "PE買入門檻(P20)", "PE位階%"]].to_string(index=False))


if __name__ == "__main__":
    main()
