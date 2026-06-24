# -*- coding: utf-8 -*-
"""
賣出訊號監看器 — 會買是徒弟,會賣才是師傅
=======================================================================
每日對 portfolio.yaml 的持倉算 6 種賣出觸發器,綜合成燈號,🟠以上 email。
資料全部自給(FinMind 股價/月營收 + 體檢總表),不另接 goodinfo/Yahoo。

6 種觸發器(從強到弱):
  🔴 ① 證偽條件   : portfolio.yaml 月營收/毛利條件達一條 → 出場(無條件)
  🔴 ② 估值過熱   : PE位階≥85 『且』(EPS成長下滑 或 PEG>2.5)→ 嚴格,不誤殺成長股
                    (循環股改看 PBR位階≥90)
  🟡 ③ 賠率反轉   : (目標-現價)/(現價-停損);>2持有 / 1~2減碼 / <1出場 / <0.5立刻
  🟡 ④ 拐點逆轉   : 含金量<0.8 或 月營收連2月動能下降 或 最新月YoY轉負
  🟢 ⑤ 部位管理   : 漲>30% 或 單一部位>20%(需填真實張數,否則略過)
  🟢 ⑥ 反向測試   : 「現價下用買進評分,我還會買嗎?」不會買=賣訊號(統一買賣邏輯)

綜合燈號:🔴出場 / 🟠減碼 / 🟡警戒 / 🟢抱緊。輸出 data/賣出訊號監看.xlsx + email。
★ 買進價/張數為草稿時,賠率/部位仍會算但僅供參考(草稿成本≈現價時賠率會失真)。
"""
import os, smtplib, ssl
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import yaml

TOKEN      = os.environ.get("FINMIND_TOKEN", "")
PORTFOLIO  = "portfolio.yaml"
HEALTH     = "data/台股_體檢總表.xlsx"
FIN_FILE   = "data/台股財報估值.xlsx"
OUT        = "data/賣出訊號監看.xlsx"
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "")
MAIL_TO    = os.environ.get("MAIL_TO", "") or GMAIL_USER


def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try: dl.login_by_token(api_token=TOKEN)
        except Exception as e: print("token失敗:", e)
    return dl


def latest_close(dl, sid):
    start = (datetime.now() - timedelta(days=20)).strftime("%Y-%m-%d")
    try:
        df = dl.taiwan_stock_daily(stock_id=sid, start_date=start)
        df = df.dropna(subset=["close"]).sort_values("date")
        return round(float(df["close"].iloc[-1]), 1) if len(df) else None
    except Exception as e:
        print(f"  {sid} 股價失敗:{e}"); return None


def recent_yoy(dl, sid, months=6):
    start = (datetime.now() - timedelta(days=550)).strftime("%Y-%m-%d")
    try:
        df = dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start)
    except Exception:
        return []
    if df is None or df.empty: return []
    df = df.sort_values("date").reset_index(drop=True)
    df["yoy"] = df["revenue"].pct_change(12) * 100
    return [round(float(x), 1) for x in df["yoy"].dropna().tail(months)]


def load_health():
    """讀體檢總表 → {代號: {...}};沒檔回空。"""
    if not os.path.exists(HEALTH): return {}
    try:
        h = pd.read_excel(HEALTH, "體檢總表"); h["代號"] = h["代號"].astype(str)
    except Exception:
        return {}
    out = {}
    for _, r in h.iterrows():
        out[r["代號"]] = {
            "評等": r.get("評等"),
            "PE位階": pd.to_numeric(r.get("PE位階"), errors="coerce"),
            "PBR位階": pd.to_numeric(r.get("PBR位階"), errors="coerce"),
            "PEG": pd.to_numeric(r.get("PEG"), errors="coerce"),
            "未來估值": r.get("未來估值"),
            "含金量": pd.to_numeric(r.get("含金量"), errors="coerce"),
            "EPS近3y%": pd.to_numeric(r.get("EPS近3y%"), errors="coerce"),
            "EPS5y%": pd.to_numeric(r.get("EPS5y%"), errors="coerce"),
            "循環": "循環" in str(r.get("循環股", "")),
            "殖利率": pd.to_numeric(r.get("殖利率"), errors="coerce"),
        }
    return out


def latest_q_gm(sid):
    try:
        q = pd.read_excel(FIN_FILE, "逐季毛利率", index_col=0)
        for k in q.index:
            if str(k).split()[0] == sid:
                v = q.loc[k].dropna()
                if len(v): return round(float(v.iloc[-1]), 2)
    except Exception: pass
    return None


def evaluate(pos, close, yoys, qgm, h):
    """回 (燈號, 各觸發器dict)。燈號:🔴出場/🟠減碼/🟡警戒/🟢抱緊。"""
    sid = pos["sid"]
    sig = {}
    cyc = h.get("循環", False) if h else False

    # ── ① 證偽條件(🔴 無條件)──
    falsify = []
    if yoys:
        neg = 0
        for v in reversed(yoys):
            if v < 0: neg += 1
            else: break
        thr_n = pos.get("證偽_月營收", {}).get("連續轉負月數", 99)
        if neg >= thr_n:
            falsify.append(f"月營收連{neg}月轉負(門檻{thr_n})")
    thr_gm = pos.get("證偽_毛利率", {}).get("跌破")
    if qgm is not None and thr_gm is not None and qgm < thr_gm:
        falsify.append(f"毛利{qgm}%跌破{thr_gm}%")
    sig["①證偽"] = "🔴 " + "、".join(falsify) if falsify else ""

    # ── ② 估值過熱(🔴 嚴格:位階高 + 成長下滑/PEG高)──
    overheat = ""
    if h:
        e3, e5, peg = h["EPS近3y%"], h["EPS5y%"], h["PEG"]
        eps_slow = pd.notna(e3) and pd.notna(e5) and e3 < e5      # 近3年成長 < 近5年 = 減速
        peg_high = pd.notna(peg) and peg > 2.5
        if cyc:    # 循環股看 PBR
            if pd.notna(h["PBR位階"]) and h["PBR位階"] >= 90:
                overheat = f"🔴 循環股PBR位階{h['PBR位階']:.0f}(≥90)"
        else:
            if pd.notna(h["PE位階"]) and h["PE位階"] >= 85 and (eps_slow or peg_high):
                why = "EPS減速" if eps_slow else f"PEG{peg:.1f}>2.5"
                overheat = f"🔴 PE位階{h['PE位階']:.0f}≥85 且 {why}"
    sig["②估值過熱"] = overheat

    # ── ③ 賠率反轉(🟡)──
    tgt, stop = pos.get("目標價"), pos.get("停損價")
    rr_txt = ""; rr = None
    if tgt and stop and close and close > stop:
        rr = (tgt - close) / (close - stop)
        if rr < 0.5:   rr_txt = f"🔴 賠率{rr:.2f}<0.5 立刻賣"
        elif rr < 1:   rr_txt = f"🟠 賠率{rr:.2f}<1 應出場"
        elif rr < 2:   rr_txt = f"🟡 賠率{rr:.2f} 考慮減碼"
        else:          rr_txt = f"🟢 賠率{rr:.2f} 持有"
    elif tgt and stop and close and close <= stop:
        rr_txt = f"🔴 跌破停損價{stop}"
    sig["③賠率"] = rr_txt
    sig["_rr"] = rr

    # ── ④ 拐點逆轉(🟡)──
    rev = []
    if h and pd.notna(h["含金量"]) and h["含金量"] < 0.8:
        rev.append(f"含金量{h['含金量']:.2f}<0.8")
    if yoys and len(yoys) >= 3:
        if yoys[-1] < 0:
            rev.append(f"最新月YoY{yoys[-1]}%轉負")
        elif yoys[-1] < yoys[-2] < yoys[-3]:
            rev.append(f"月營收動能連2月降({yoys[-3]}→{yoys[-2]}→{yoys[-1]})")
    sig["④拐點逆轉"] = ("🟡 " + "、".join(rev)) if rev else ""

    # ── ⑤ 部位管理(🟢,需真實成本/張數)──
    pos_mgmt = ""
    buy = pos.get("買進", {}).get("價")
    shares = pos.get("張數", 0)
    if buy and close and shares and shares > 0:
        ret = (close / buy - 1) * 100
        if ret > 30:
            pos_mgmt = f"🟢 報酬+{ret:.0f}%>30%,可部分了結"
    sig["⑤部位"] = pos_mgmt

    # ── ⑥ 反向測試:現價下還會買嗎?(🟢,統一買賣邏輯)──
    # 買進門檻:評等A/B + 估值不過熱 + 賠率≥2 + 含金量≥0.8(循環股放寬看PBR)
    rebuy = ""
    if h:
        ok_grade = str(h["評等"]) in ("A", "B")
        ok_cash = pd.notna(h["含金量"]) and h["含金量"] >= 0.8
        ok_rr = (rr is None) or (rr >= 2)
        ok_val = not overheat
        if not (ok_grade and ok_cash and ok_rr and ok_val):
            fails = []
            if not ok_grade: fails.append("非A/B級")
            if not ok_cash:  fails.append("含金量弱")
            if not ok_rr:    fails.append("賠率<2")
            if not ok_val:   fails.append("估值過熱")
            rebuy = "🟢 現價下不會買(" + "、".join(fails) + ")"
    sig["⑥反向測試"] = rebuy

    # ── 綜合燈號 ──
    if sig["①證偽"] or sig["②估值過熱"] or (rr is not None and rr < 1):
        light = "🔴 出場"
    elif (rr is not None and rr < 2) or sig["④拐點逆轉"] or sig["⑤部位"]:
        light = "🟠 減碼"
    elif sig["⑥反向測試"]:
        light = "🟡 警戒"
    else:
        light = "🟢 抱緊"
    return light, sig


def send_email(subject, html):
    if not (GMAIL_USER and GMAIL_PASS and MAIL_TO):
        print("未設 GMAIL,略過寄信"); return
    msg = MIMEMultipart("alternative")
    msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, MAIL_TO
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
            s.starttls(context=ssl.create_default_context())
            s.login(GMAIL_USER, GMAIL_PASS); s.send_message(msg)
        print("已寄出 →", MAIL_TO)
    except Exception as e:
        print("寄信失敗:", e)


def main():
    cfg = yaml.safe_load(open(PORTFOLIO, encoding="utf-8"))
    positions = cfg.get("positions", [])
    if not positions:
        print("portfolio.yaml 無持倉"); return
    dl = make_loader()
    health = load_health()
    print(f"監看 {len(positions)} 檔持倉,體檢覆蓋 {len(health)} 檔")

    rows = []
    for pos in positions:
        sid = pos["sid"]
        close = latest_close(dl, sid)
        yoys = recent_yoy(dl, sid)
        qgm = latest_q_gm(sid)
        h = health.get(sid)
        light, sig = evaluate(pos, close, yoys, qgm, h)
        rows.append({
            "代號": sid, "名稱": pos["name"], "燈號": light,
            "現價": close, "成本(草稿)": pos.get("買進", {}).get("價"),
            "目標": pos.get("目標價"), "停損": pos.get("停損價"),
            "①證偽": sig["①證偽"], "②估值過熱": sig["②估值過熱"],
            "③賠率": sig["③賠率"], "④拐點逆轉": sig["④拐點逆轉"],
            "⑤部位": sig["⑤部位"], "⑥反向測試": sig["⑥反向測試"],
        })
        print(f"  {sid} {pos['name']:6s} {light}  {sig['③賠率']}")

    df = pd.DataFrame(rows)
    order = {"🔴 出場": 0, "🟠 減碼": 1, "🟡 警戒": 2, "🟢 抱緊": 3}
    df["_o"] = df["燈號"].map(lambda x: order.get(x, 9))
    df = df.sort_values("_o").drop(columns="_o")
    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="賣出訊號", index=False)

    today = datetime.now().strftime("%Y-%m-%d")
    act = df[df["燈號"].isin(["🔴 出場", "🟠 減碼"])]
    n_red = (df["燈號"] == "🔴 出場").sum()
    n_org = (df["燈號"] == "🟠 減碼").sum()
    print(f"\n🔴出場 {n_red} / 🟠減碼 {n_org} / 🟡警戒 {(df['燈號']=='🟡 警戒').sum()} / 🟢抱緊 {(df['燈號']=='🟢 抱緊').sum()}")

    if len(act):
        cols = ["代號","名稱","燈號","現價","目標","停損","①證偽","②估值過熱","③賠率","④拐點逆轉"]
        html = (f"<h3>{today} 賣出訊號(🔴出場 {n_red} / 🟠減碼 {n_org})</h3>"
                + act[cols].to_html(index=False, border=1)
                + "<p style='color:#888'>說明:燈號=六觸發器綜合。🔴=證偽/估值過熱/賠率<1;"
                  "🟠=賠率1~2/拐點逆轉/部位過大。買進價為草稿時賠率僅參考。</p>")
        send_email(f"【賣出訊號】{today} 🔴{n_red}/🟠{n_org} 檔", html)
    else:
        print("今日無 🔴/🟠 訊號,不寄信")
    print(f"完成 → {OUT}")


if __name__ == "__main__":
    main()
