# -*- coding: utf-8 -*-
"""
每日 PE 買入區間監看器 v2(自算 PER + 本益比河流圖)
=====================================================
為什麼改寫:FinMind 的 taiwan_stock_per_pbr「PER」欄位不可靠(基準 EPS 不一致、
未還原股本變動、且與 Goodinfo/財報狗對不上,例:奇鋐顯示 12 但實際 ~47)。
改成自算,與 Goodinfo 本益比河流圖同一套:

    PER = 收盤價 ÷ 近四季EPS
    近四季EPS = 最近 4 個單季 EPS 加總(FinMind 財報,已驗證準確)

買入區間:今日「自算 PER」≤ 該股近5年「自算 PER」的 P20(PE 位階 ≤ BUY_PCTL%)。
同時輸出本益比河流圖:目前近四季EPS × {10,14,18,22,26,30} 倍對應價格,看股價落在哪一帶。

資料來源(每檔 3 次呼叫,仍輕):
  taiwan_stock_daily(收盤價)、taiwan_stock_financial_statement(單季EPS)、
  taiwan_stock_per_pbr(只取 dividend_yield 殖利率,不用其 PER)。
清單:tickers_watch.txt。輸出:data/PE買入區間監看.xlsx。有訊號寄 Gmail。

★ 寄信需設 Secrets:GMAIL_USER、GMAIL_APP_PASSWORD(Google 應用程式密碼);可選 MAIL_TO。
"""

import os, time, smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pandas as pd
import requests

# ---------- 設定 ----------
TOKEN       = os.environ.get("FINMIND_TOKEN", "")
WATCH_FILE  = "tickers_watch.txt"
HEALTH_FILE = "data/台股_體檢總表.xlsx"   # 體檢評分(品質濾鏡來源);無此檔則退回裸PE監看
OUTPUT      = "data/PE買入區間監看.xlsx"
BUY_PCTL    = 20                 # 買入區間:自算 PE 位階 ≤ 此百分位
YEARS       = 5                  # PE 位階用近幾年分布
MIN_DAYS    = 250               # PER 序列至少要有的天數
REQ_SLEEP   = 0.3
BANDS       = [10, 14, 18, 22, 26, 30]   # 本益比河流圖倍數

GMAIL_USER  = os.environ.get("GMAIL_USER", "")
GMAIL_PASS  = os.environ.get("GMAIL_APP_PASSWORD", "")
MAIL_TO     = os.environ.get("MAIL_TO", "") or GMAIL_USER


# ---------- 清單 / FinMind ----------
def load_watch():
    out = []
    if not os.path.exists(WATCH_FILE):
        return out
    for line in open(WATCH_FILE, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        for tok in line.replace(",", " ").replace("\t", " ").split():
            if tok.strip().isdigit():
                out.append(tok.strip())
    return list(dict.fromkeys(out))

def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try:
            dl.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("token 登入失敗(改用免費額度):", e)
    return dl

def _is_rate_limit(e):
    msg = str(e).lower()
    return any(k in msg for k in ("limit", "402", "429", "too many", "exceed", "request"))

def seconds_to_next_hour(buffer=45):
    """距下一個整點還有幾秒(FinMind 額度每小時重置),多加 buffer 秒保險。"""
    now = datetime.now()
    nxt = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    return max(5, int((nxt - now).total_seconds()) + buffer)

def names_map(dl):
    try:
        info = dl.taiwan_stock_info()
        return {str(r["stock_id"]): str(r["stock_name"]) for _, r in info.iterrows()}
    except Exception:
        return {}

def get_dividend_yield(dl, sid, start):
    """只取殖利率(per_pbr 的 dividend_yield 是價格基礎,可靠;不用它的 PER)。"""
    try:
        df = dl.taiwan_stock_per_pbr(stock_id=sid, start_date=start)
    except Exception:
        h = {"Authorization": f"Bearer {TOKEN}"} if TOKEN else {}
        r = requests.get("https://api.finmindtrade.com/api/v4/data",
                         params={"dataset": "TaiwanStockPER", "data_id": sid, "start_date": start},
                         headers=h, timeout=20)
        df = pd.DataFrame(r.json().get("data", []))
    if df is None or df.empty or "dividend_yield" not in df.columns:
        return None
    dy = pd.to_numeric(df.sort_values("date")["dividend_yield"], errors="coerce").dropna()
    return round(float(dy.iloc[-1]), 2) if len(dy) else None


# ---------- 近四季EPS / 自算 PER ----------
def pivot(df):
    if df is None or df.empty or "type" not in df.columns:
        return pd.DataFrame()
    return df.pivot_table(index="date", columns="type", values="value", aggfunc="first").sort_index()

def ttm_eps(fin):
    """單季 EPS → 近四季EPS 序列(index=季底日,加上公布落後天數當『生效日』,避免未卜先知)。"""
    piv = pivot(fin)
    if piv.empty or "EPS" not in piv.columns:
        return None
    eps = pd.to_numeric(piv["EPS"], errors="coerce").dropna()
    if len(eps) < 4:
        return None
    ttm = eps.rolling(4).sum().dropna()
    rows = []
    for d, v in ttm.items():
        qend = pd.to_datetime(d)
        lag = 90 if qend.month == 12 else 45        # Q4(年報)約90天、其餘約45天才公布
        rows.append((qend + timedelta(days=lag), float(v)))
    return pd.DataFrame(rows, columns=["生效日", "近四季EPS"]).sort_values("生效日")

def analyze(price_df, fin_df, dy):
    if price_df is None or price_df.empty or "close" not in price_df.columns:
        return None
    tt = ttm_eps(fin_df)
    if tt is None or tt.empty:
        return None
    p = price_df[["date", "close"]].copy()
    p["close"] = pd.to_numeric(p["close"], errors="coerce")
    p = p.dropna().sort_values("date")
    p["生效日"] = pd.to_datetime(p["date"])
    m = pd.merge_asof(p.sort_values("生效日"), tt, on="生效日", direction="backward")
    m = m[(m["近四季EPS"] > 0)].copy()
    if m.empty:
        return None
    m["PER"] = m["close"] / m["近四季EPS"]
    s = m["PER"].replace([float("inf")], pd.NA).dropna()
    s = s[s > 0]
    if len(s) < MIN_DAYS:
        return None
    cur     = float(s.iloc[-1])
    close   = float(m["close"].iloc[-1])
    eps_ttm = float(m["近四季EPS"].iloc[-1])
    p20     = float(s.quantile(BUY_PCTL / 100))
    rank    = round(float((s <= cur).mean() * 100))
    out = {
        "收盤": round(close, 1),
        "近四季EPS": round(eps_ttm, 2),
        "PER現(自算)": round(cur, 2),
        "PE買入門檻(P20)": round(p20, 2),
        "PE位階%": rank,
        "買入價(P20×EPS)": round(p20 * eps_ttm, 1),
        "距買入區間%": round((cur / p20 - 1) * 100, 1),
        "PE5年低": round(float(s.min()), 2),
        "PE5年高": round(float(s.max()), 2),
        "殖利率%": dy,
        "最新日": str(m["date"].iloc[-1]),
        "_buy": rank <= BUY_PCTL,
    }
    for x in BANDS:                                  # 本益比河流圖各倍數對應價
        out[f"{x}x"] = round(eps_ttm * x, 1)
    return out


# ---------- 體檢分數(品質濾鏡) ----------
def load_health():
    """讀體檢總表 → {代號: {評等,品質總分,含金量,EPS3y,循環,PBR位階}}。無檔則回空 dict(退回裸PE)。"""
    if not os.path.exists(HEALTH_FILE):
        print("無體檢總表,退回裸 PE 監看(建議先跑 stock_health_check.py)"); return {}
    try:
        h = pd.read_excel(HEALTH_FILE, "體檢總表"); h["代號"] = h["代號"].astype(str)
    except Exception as e:
        print("讀體檢總表失敗,退回裸PE:", e); return {}
    out = {}
    for _, r in h.iterrows():
        out[r["代號"]] = {
            "評等": r.get("評等"), "品質總分": r.get("品質總分"),
            "含金量": pd.to_numeric(r.get("含金量"), errors="coerce"),
            "EPS3y": pd.to_numeric(r.get("EPS近3y%"), errors="coerce"),
            "PBR位階": pd.to_numeric(r.get("PBR位階"), errors="coerce"),
            "循環": "循環" in str(r.get("循環股", "")),
        }
    return out


def classify(r, h):
    """三層訊號:⭐優質買點 / ⚠️便宜陷阱 / 🔄循環買點(PBR) / (空=非買點)。
    r=今日PE分析(含PE位階%);h=該檔體檢(可能None)。"""
    pe_buy = r["PE位階%"] <= BUY_PCTL
    if not h:                                     # 無體檢資料
        return "❔未評(無體檢)" if pe_buy else ""
    if h["評等"] == "金融🏦":                       # 金融股:不評分,看PBR/殖利率自行判斷
        return "🏦金融(看PBR/殖利率)" if pe_buy else ""
    grade = h["評等"]; g = h["含金量"]; e3 = h["EPS3y"]
    if h["循環"]:                                  # 循環股:PER失真,改看PBR位階
        pbr = h["PBR位階"]
        if pd.notna(pbr) and pbr <= BUY_PCTL:
            # 循環買點也要品質gate:健康的循環(非D級+有現金)才算,否則是陷阱
            healthy = (grade != "D") and (pd.notna(g) and g >= 0.8)
            return "🔄循環買點(PBR)" if healthy else "⚠️便宜陷阱"
        return ""
    if not pe_buy:                                 # 非循環:沒落入PE買區就不是買點
        return ""
    cash_ok = (pd.notna(g) and g >= 1.0) and (pd.notna(e3) and e3 > 0)
    if grade == "A" and cash_ok:
        return "⭐A級買點"                          # 便宜+頂級品質+真成長
    if grade == "B" and cash_ok:
        return "✅B級買點"                          # 便宜+合格(成長普通,偏防禦/領息)
    return "⚠️便宜陷阱"


# ---------- Email ----------
def send_email(subject, html):
    if not (GMAIL_USER and GMAIL_PASS and MAIL_TO):
        print("未設 GMAIL_USER/GMAIL_APP_PASSWORD,略過寄信"); return
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, MAIL_TO
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as srv:
            srv.starttls(context=ssl.create_default_context())
            srv.login(GMAIL_USER, GMAIL_PASS)
            srv.send_message(msg)
        print("已寄出 email →", MAIL_TO)
    except Exception as e:
        print("寄信失敗:", e)


# ---------- 主流程 ----------
def main():
    dl = make_loader()
    nm = names_map(dl)
    y = datetime.now().year
    p_start = f"{y - YEARS}-01-01"          # 股價近5年
    f_start = f"{y - YEARS - 2}-01-01"      # 財報多抓2年,確保序列起點就有近四季EPS
    watch = load_watch()
    health = load_health()
    print(f"監看 {len(watch)} 檔,買入區間 = 自算 PE 位階 ≤ {BUY_PCTL}%;體檢覆蓋 {len(health)} 檔")
    rows = []
    for i, sid in enumerate(watch, 1):
        r = None
        while True:                                    # 撞額度就睡到整點再續、不跳過此檔
            try:
                price = dl.taiwan_stock_daily(stock_id=sid, start_date=p_start)
                fin   = dl.taiwan_stock_financial_statement(stock_id=sid, start_date=f_start)
                dy    = get_dividend_yield(dl, sid, p_start)
                r = analyze(price, fin, dy)
                break
            except Exception as e:
                if _is_rate_limit(e):
                    wait = seconds_to_next_hour()
                    print(f"  ⏸ 疑似 FinMind 額度用罄 → 睡 {wait//60} 分到整點再續(不跳過 {sid})")
                    time.sleep(wait); continue
                print(f"  ! {sid} 失敗:{e}"); r = None; break
        if r:
            h = health.get(sid)
            sig = classify(r, h)
            row = {"代號": sid, "名稱": nm.get(sid, sid), **r}
            row["訊號"] = sig
            row["評等"] = h["評等"] if h else ""
            row["含金量"] = h["含金量"] if h else None
            row["EPS近3y%"] = h["EPS3y"] if h else None
            row["PBR位階"] = h["PBR位階"] if h else None
            row["循環"] = "⚠️" if (h and h["循環"]) else ""
            rows.append(row)
            print(f"[{i}/{len(watch)}] {sid} {nm.get(sid,sid):6s} 位階{r['PE位階%']:>3} {sig or ''}")
        time.sleep(REQ_SLEEP)

    if not rows:
        print("無有效資料"); return
    df = pd.DataFrame(rows).drop(columns="_buy").sort_values("PE位階%")

    star = df[df["訊號"] == "⭐A級買點"]
    bgrade = df[df["訊號"] == "✅B級買點"]
    trap = df[df["訊號"] == "⚠️便宜陷阱"]
    cyc  = df[df["訊號"] == "🔄循環買點(PBR)"]
    fin  = df[df["訊號"] == "🏦金融(看PBR/殖利率)"]
    unrated = df[df["訊號"] == "❔未評(無體檢)"]

    view = ["代號", "名稱", "訊號", "評等", "收盤", "PER現(自算)", "PE位階%", "PBR位階",
            "含金量", "EPS近3y%", "循環", "買入價(P20×EPS)", "距買入區間%", "殖利率%"]
    def v(d):
        return d[[c for c in view if c in d.columns]]

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        v(star).to_excel(xw, sheet_name="A級買點", index=False)
        v(bgrade).to_excel(xw, sheet_name="B級買點", index=False)
        v(cyc).to_excel(xw, sheet_name="循環股買點(看PBR)", index=False)
        v(fin).to_excel(xw, sheet_name="金融(看PBR殖利率)", index=False)
        v(trap).to_excel(xw, sheet_name="便宜但陷阱", index=False)
        df.to_excel(xw, sheet_name="全部監看", index=False)

    today = datetime.now().strftime("%Y-%m-%d")
    print(f"\n⭐A級買點 {len(star)} / ✅B級買點 {len(bgrade)} / 🔄循環 {len(cyc)} / "
          f"🏦金融 {len(fin)} / ⚠️陷阱 {len(trap)} / ❔未評 {len(unrated)}")
    # 寄 A 級 + B 級買點(真訊號,分級呈現);其餘只附摘要
    if len(star) or len(bgrade):
        cols = ["代號", "名稱", "評等", "收盤", "PER現(自算)", "PE位階%", "含金量", "EPS近3y%", "買入價(P20×EPS)", "殖利率%"]
        html = f"<h3>{today} 買點訊號(體檢加持)</h3>"
        if len(star):
            html += f"<h4>⭐ A級買點 {len(star)} 檔(便宜+頂級品質+真成長)</h4>" + star[cols].to_html(index=False, border=1)
        if len(bgrade):
            html += f"<h4>✅ B級買點 {len(bgrade)} 檔(便宜+合格,成長普通偏防禦/領息)</h4>" + bgrade[cols].to_html(index=False, border=1)
        html += (f"<p>另:🔄循環買點 {len(cyc)}、🏦金融 {len(fin)}、⚠️便宜陷阱 {len(trap)}(已過濾),詳見 {OUTPUT}。</p>")
        send_email(f"【買點】{today} A級{len(star)}/B級{len(bgrade)} 檔(體檢加持)", html)
    else:
        print("今日無 A/B 級買點,不寄信")

    print(f"\n完成 → {OUTPUT};今日買入區間 {len(buy)} 檔 / 監看 {len(allv)} 檔")
    if len(buy):
        print(buy[["代號", "名稱", "收盤", "PER現(自算)", "PE買入門檻(P20)", "PE位階%"]].to_string(index=False))


if __name__ == "__main__":
    main()
