# -*- coding: utf-8 -*-
"""
美股翻轉即時警報 us_flip_alert.py
=======================================================================
每工作日跑, 對 watchlist_us.txt 每檔算:
  - 內部人 4Q 買賣比 (insider-trading/statistics)
  - 國會交易 90d 淨買賣 (senate + house)
  - 國會強訊號 (>= 3 同向 3:1)

對比自己昨日 snapshot, 偵測翻轉:
  🔴 內部人 4Q 買賣比 由 > 1 → < 0.5 (淨買轉大賣)
  🟢 內部人 4Q 買賣比 由 < 0.3 → > 1 (大賣轉淨買)
  🔴 國會 90d 淨買 由 > 3 → < -1
  🟢 國會 90d 淨賣 由 < -3 → > 1
  🚨 國會強訊號翻轉 (強買↔強賣)

只在有翻轉時寄 email

輸出:
  data/美股_翻轉警報.xlsx
  /tmp/us_flip_subject.txt + body.html
"""
import os, json, sys, time, requests
import pandas as pd
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

KEY = os.environ.get("FMP_API_KEY", "")
BASE = "https://financialmodelingprep.com/stable"
WATCHLIST_FILE = os.environ.get("WATCHLIST_FILE", "data/watchlist_us.txt")
WORKERS = int(os.environ.get("WORKERS", "6"))
DST = "data/美股_翻轉警報.xlsx"
PREV_FILE = "data/us_flip_alert_prev.json"

TPE = timezone(timedelta(hours=8))
TODAY = datetime.now(TPE).strftime("%Y-%m-%d")

# 翻轉閾值
INSIDER_FLIP_HI = 1.0   # 高於此為「內部人淨買」
INSIDER_FLIP_LO = 0.3   # 低於此為「內部人大賣」
CONGRESS_STRONG_NET = 3 # 90d 淨買賣 >= 3 才算強訊號


def load_watchlist():
    env = os.environ.get("TICKERS", "").strip()
    if env:
        toks = [t.strip().upper() for t in env.replace(",", " ").split() if t.strip()]
        return list(dict.fromkeys([t for t in toks if t and not t.startswith("#")]))
    if os.path.exists(WATCHLIST_FILE):
        toks = []
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.split("#", 1)[0].strip()
                if not line: continue
                toks.extend(t.strip().upper() for t in line.split() if t.strip())
        return list(dict.fromkeys(toks))
    return "NVDA AVGO MSFT".split()


def get(endpoint, **params):
    if not KEY: return None
    params["apikey"] = KEY
    for _ in range(3):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=15)
            if r.status_code == 429: time.sleep(2); continue
            if r.status_code != 200: return None
            return r.json()
        except Exception: time.sleep(1)
    return None


def compute(sym):
    """對一檔算內部人 4Q + 國會 90d"""
    out = {"代號": sym}

    # 內部人 4Q 合計
    ins = get("insider-trading/statistics", symbol=sym) or []
    if isinstance(ins, list) and ins:
        recent4 = ins[:4]
        t_acq = sum((r.get("totalAcquired") or 0) for r in recent4)
        t_dis = sum((r.get("totalDisposed") or 0) for r in recent4)
        n_acq = sum((r.get("acquiredTransactions") or 0) for r in recent4)
        n_dis = sum((r.get("disposedTransactions") or 0) for r in recent4)
        out["內部人4Q買量"] = t_acq
        out["內部人4Q賣量"] = t_dis
        out["內部人4Q買筆"] = n_acq
        out["內部人4Q賣筆"] = n_dis
        if t_dis > 0:
            out["內部人買賣比"] = round(t_acq / t_dis, 3)
        elif t_acq > 0:
            out["內部人買賣比"] = 99.0
        else:
            out["內部人買賣比"] = None

    # 國會 90d
    sen = get("senate-trades", symbol=sym) or []
    hou = get("house-trades", symbol=sym) or []
    cutoff = (datetime.now(TPE) - timedelta(days=90)).strftime("%Y-%m-%d")
    all_recent = [t for t in (sen if isinstance(sen, list) else [])
                  if t.get("transactionDate", "") >= cutoff]
    all_recent += [t for t in (hou if isinstance(hou, list) else [])
                    if t.get("transactionDate", "") >= cutoff]
    g_buy = sum(1 for t in all_recent if t.get("type","").lower() in ("purchase","buy"))
    g_sell = sum(1 for t in all_recent if t.get("type","").lower() in ("sale","sell"))
    out["國會90d買"] = g_buy
    out["國會90d賣"] = g_sell
    # 強訊號
    if g_buy >= 3 and g_buy >= g_sell * 3:
        out["國會強訊號"] = "強買"
    elif g_sell >= 3 and g_sell >= g_buy * 3:
        out["國會強訊號"] = "強賣"
    else:
        out["國會強訊號"] = None

    return sym, out


def detect_flips(cur_map, prev_map):
    """偵測翻轉 (絕對狀態 + delta)"""
    flips = []
    for sym, cur in cur_map.items():
        events = []; severity = 0

        # ═══ 絕對狀態訊號 (首次跑也能觸發) ═══
        cur_r = cur.get("內部人買賣比")
        cur_sig = cur.get("國會強訊號")
        cur_buy = cur.get("國會90d買", 0) or 0
        cur_sell = cur.get("國會90d賣", 0) or 0
        cur_net = cur_buy - cur_sell

        # 1. 極端內部人單邊 (無需國會)
        if cur_r is not None:
            if cur_r < 0.05:  # 內部人 4Q 買量幾乎為 0
                events.append(f"🔴 極端內部人賣壓(4Q 買賣比 {cur_r})")
                severity += 2
            elif cur_r > 5:  # 內部人瘋狂淨買
                events.append(f"🟢 極端內部人淨買(4Q 買賣比 {cur_r})")
                severity += 2

        # 2. 國會強訊號 (無需內部人)
        if cur_sig == "強賣":
            events.append(f"🔴 國會強賣({cur_buy}買/{cur_sell}賣)")
            severity += 2
        elif cur_sig == "強買":
            events.append(f"🟢 國會強買({cur_buy}買/{cur_sell}賣)")
            severity += 2

        # 3. 內外雙賣 (bonus 加分)
        if cur_r is not None and cur_r < 0.15 and cur_net <= -2:
            events.append(f"🚨 內外一致減碼")
            severity += 3
        # 內外雙買
        elif cur_r is not None and cur_r > 2 and cur_net >= 2:
            events.append(f"🚨 內外一致加碼")
            severity += 3

        # ═══ Delta 訊號 (需要 prev) ═══
        prev = prev_map.get(sym)
        if prev:
            prv_r = prev.get("內部人買賣比")
            # 內部人 4Q 買賣比 flip
            if cur_r is not None and prv_r is not None:
                if prv_r >= INSIDER_FLIP_HI and cur_r < INSIDER_FLIP_LO:
                    events.append(f"🔴 內部人翻賣 ({prv_r} → {cur_r})")
                    severity += 3
                elif prv_r < INSIDER_FLIP_LO and cur_r >= INSIDER_FLIP_HI:
                    events.append(f"🟢 內部人翻買 ({prv_r} → {cur_r})")
                    severity += 3
                elif prv_r >= 2 and cur_r < 1:
                    events.append(f"🟠 內部人買氣減弱 ({prv_r} → {cur_r})")
                    severity += 1

            # 國會 90d 淨買賣 flip
            prv_net = (prev.get("國會90d買", 0) or 0) - (prev.get("國會90d賣", 0) or 0)
            if prv_net >= CONGRESS_STRONG_NET and cur_net < -1:
                events.append(f"🔴 國會翻賣 (淨 {prv_net:+d} → {cur_net:+d})")
                severity += 2
            elif prv_net < -CONGRESS_STRONG_NET and cur_net > 1:
                events.append(f"🟢 國會翻買 (淨 {prv_net:+d} → {cur_net:+d})")
                severity += 2

            # 國會強訊號翻轉
            prv_sig = prev.get("國會強訊號")
            if prv_sig == "強買" and cur_sig == "強賣":
                events.append("🚨 國會強買→強賣")
                severity += 3
            elif prv_sig == "強賣" and cur_sig == "強買":
                events.append("🚨 國會強賣→強買")
                severity += 3
            elif prv_sig != cur_sig and cur_sig in ("強買","強賣"):
                events.append(f"🆕 國會新出現 {cur_sig}")
                severity += 1

        if events:
            flips.append({
                "代號": sym, "訊號": " / ".join(events), "嚴重度": severity,
                "內部人買賣比(今)": cur_r,
                "內部人買賣比(昨)": (prev or {}).get("內部人買賣比"),
                "國會90d買": cur.get("國會90d買"),
                "國會90d賣": cur.get("國會90d賣"),
                "國會強訊號(今)": cur_sig,
                "國會強訊號(昨)": (prev or {}).get("國會強訊號"),
            })
    return sorted(flips, key=lambda x: -x["嚴重度"])


def build_email(flips, six_map=None):
    if not flips: return None, None
    n_severe = sum(1 for f in flips if f["嚴重度"] >= 3)
    n_normal = len(flips) - n_severe
    subject = f"🚨 [US 翻轉 {TODAY}] {len(flips)} 檔 (嚴重 {n_severe})"

    body = [f"""<html><body style='font-family:-apple-system,sans-serif;max-width:1100px'>
<h2>🚨 美股 內部人 + 國會 翻轉警報 — {TODAY}</h2>
<p style='color:#666'>每工作日跑, 只有翻轉時才寄</p>
<p><b>{len(flips)}</b> 檔翻轉 | 🚨 嚴重 <b>{n_severe}</b> | 一般 {n_normal}</p>

<div style='background:#f5f5ff;padding:12px;border-left:4px solid #55b;margin:12px 0;font-size:13px'>
<h4 style='margin:0 0 8px 0'>⏰ 資料時效性重要提醒</h4>
<table style='border-collapse:collapse;width:100%;font-size:12px'>
<tr style='background:#eee'><th style='padding:4px'>資料</th><th>法定申報</th><th>實際延遲</th><th>時效性</th></tr>
<tr><td style='padding:4px'><b>內部人 Form 4</b></td><td>交易後 2 個工作日</td><td>1-3 天</td><td>⚡ <b>接近即時</b></td></tr>
<tr><td style='padding:4px'><b>國會 PTR</b></td><td>交易後 <b>45 天</b></td><td>平均 20-45 天</td><td>🐌 <b>看到時已 2-3 個月前</b></td></tr>
</table>
</div>

<div style='background:#fffbe0;padding:12px;border-left:4px solid #fa0;margin:12px 0;font-size:13px'>
<h4 style='margin:0 0 8px 0'>💡 實務用法</h4>
<p style='margin:4px 0'><b>👤 內部人訊號 → 短中期決策</b></p>
<ul style='margin:2px 0;padding-left:20px'>
<li>4Q 買賣比 翻紅(1.5 → 0.05) = <b>即時警訊</b> (2-3 天內反應)</li>
<li>適合: 決定「這週該不該減碼」</li>
</ul>
<p style='margin:8px 0 4px 0'><b>🏛️ 國會訊號 → 政策風向</b></p>
<ul style='margin:2px 0;padding-left:20px'>
<li>一群政要同時買/賣同產業 = <b>政策風向球</b> (即使有 45 天延遲)</li>
<li>適合: 看法規、貿易政策、國防合約風向</li>
<li>❌ <b>不適合</b>: 當單一買賣時點決策 (訊號已延遲)</li>
</ul>
</div>

<h3 style='margin-top:20px'>⚡ 翻轉清單</h3>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:13px;width:100%'>
<tr style='background:#f0f0f0'>
<th>嚴重</th><th>代號</th><th>六維</th>
<th>訊號</th>
<th>內買賣比 今</th><th>昨</th>
<th>國會 買/賣</th>
<th>國會訊號 昨→今</th>
</tr>"""]

    for f in flips:
        code = str(f["代號"])
        severity_str = "🔴🔴🔴" if f["嚴重度"] >= 3 else ("🔴🔴" if f["嚴重度"] == 2 else "🔴")
        bg = "#ffcccc" if f["嚴重度"] >= 3 else ("#ffe5cc" if f["嚴重度"] == 2 else "")
        row_style = f"background:{bg}" if bg else ""
        six_info = six_map.get(code, "") if six_map else ""
        cr = f["內部人買賣比(今)"]
        pr = f["內部人買賣比(昨)"]
        cong_bs = f"{f.get('國會90d買','—')} / {f.get('國會90d賣','—')}"
        cong_sig = f"{f.get('國會強訊號(昨)','—')} → {f.get('國會強訊號(今)','—')}"

        body.append(f"<tr style='{row_style}'>"
                    f"<td style='text-align:center'>{severity_str}</td>"
                    f"<td>{code}</td><td>{six_info}</td>"
                    f"<td><b>{f['訊號']}</b></td>"
                    f"<td style='text-align:right'>{cr if cr is not None else '—'}</td>"
                    f"<td style='text-align:right;color:#888'>{pr if pr is not None else '—'}</td>"
                    f"<td style='text-align:center'>{cong_bs}</td>"
                    f"<td style='text-align:center'>{cong_sig}</td></tr>")
    body.append("</table>")

    body.append("""
<hr>
<h4>📌 訊號規則</h4>
<ul style='font-size:12px;color:#666'>
<li>🔴 內部人翻賣: 4Q 買賣比 由 >= 1 → < 0.3 (嚴重度 +3)</li>
<li>🟢 內部人翻買: 4Q 買賣比 由 < 0.3 → >= 1 (嚴重度 +3)</li>
<li>🟠 買氣減弱: 4Q 買賣比 由 >= 2 → < 1 (嚴重度 +1)</li>
<li>🔴🟢 國會翻賣/翻買: 90d 淨買賣 反轉 (嚴重度 +2)</li>
<li>🚨 國會強買↔強賣: 最強反轉訊號 (嚴重度 +3)</li>
<li>🆕 國會新出現強訊號: 首次觸發 (嚴重度 +1)</li>
</ul>
<p style='color:#888;font-size:11px'>資料源: FMP • 觀察 watchlist_us.txt</p>
</body></html>""")

    return subject, "\n".join(body)


def load_prev():
    if not os.path.exists(PREV_FILE): return {}
    try:
        with open(PREV_FILE, encoding="utf-8") as f: return json.load(f)
    except: return {}


def save_prev(data):
    os.makedirs(os.path.dirname(PREV_FILE), exist_ok=True)
    with open(PREV_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def main():
    if not KEY: print("⚠️ 需 FMP_API_KEY"); sys.exit(1)

    codes = load_watchlist()
    print(f"US 翻轉警報 — {len(codes)} 檔 (平行 {WORKERS})")

    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(compute, s): s for s in codes}
        done = 0
        for fut in as_completed(futs):
            sym, d = fut.result()
            done += 1
            if d: results[sym] = d
            if done % 20 == 0: print(f"  [{done}/{len(codes)}]")

    if not results:
        print("⚠️ 沒抓到資料"); sys.exit(1)

    # 六維(若有)
    six_map = {}
    six_src = "data/美股95檔_六維交叉.xlsx"
    if os.path.exists(six_src):
        try:
            six = pd.read_excel(six_src)
            six["代號"] = six["代號"].astype(str)
            for _, r in six.iterrows():
                six_map[r["代號"]] = f"{r.get('評等','')} 分{r.get('六維分','')}"
        except: pass

    # 診斷: 印分布給 user 看實際數值
    ratios = [r.get("內部人買賣比") for r in results.values() if r.get("內部人買賣比") is not None]
    congress_sigs = [r.get("國會強訊號") for r in results.values() if r.get("國會強訊號")]
    if ratios:
        ratios.sort()
        print(f"📊 內部人買賣比分布: min={ratios[0]:.2f} / 25%={ratios[len(ratios)//4]:.2f} "
              f"/ 中位={ratios[len(ratios)//2]:.2f} / max={ratios[-1]:.2f} (共 {len(ratios)} 檔有資料)")
        print(f"   < 0.05 極端賣: {sum(1 for r in ratios if r < 0.05)} 檔")
        print(f"   < 0.15 低: {sum(1 for r in ratios if r < 0.15)} 檔")
        print(f"   > 5 極端買: {sum(1 for r in ratios if r > 5)} 檔")
    print(f"📊 國會強訊號: {congress_sigs.count('強買')} 強買 / {congress_sigs.count('強賣')} 強賣")

    # 偵測翻轉
    prev = load_prev()
    prev_map = prev.get("map", {})
    flips = detect_flips(results, prev_map)
    print(f"翻轉檔數: {len(flips)}")

    # 存今日 snapshot
    save_prev({"date": TODAY, "map": results})

    if not flips:
        print("  無翻轉,不寄 email")
        with open("/tmp/us_flip_subject.txt", "w", encoding="utf-8") as f: f.write("")
        with open("/tmp/us_flip_body.html", "w", encoding="utf-8") as f: f.write("")
        return

    # 存 xlsx
    fdf = pd.DataFrame(flips)
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    fdf.to_excel(DST, index=False)
    print(f"→ {DST}")

    subject, body = build_email(flips, six_map)
    with open("/tmp/us_flip_subject.txt", "w", encoding="utf-8") as f: f.write(subject)
    with open("/tmp/us_flip_body.html", "w", encoding="utf-8") as f: f.write(body)
    print(f"→ {subject}")


if __name__ == "__main__":
    main()
