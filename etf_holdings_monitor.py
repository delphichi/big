#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主動式 ETF 持股監看  etf_holdings_monitor.py
=====================================================================
跨投信抓「主動式 ETF」每日成分,diff 昨天 → 看經理人在買賣什麼。
主動式 ETF = 真人經理人選股,持股變化是「法人決策」,訊號遠強於被動 ETF。

資料源(本機被 proxy 擋,需在 CI 跑):
  復華 fhtrust : GET /api/assetsExcel/{etf}/{YYYYMMDD}      (Excel)
  統一 ezmoney : GET /ETF/Transaction/PCFExcelNPOI?fundCode={fc}&date={ROC}&specificDate=true

每檔:存日期快照 data/etf_holdings/{code}_{YYYYMMDD}.csv → diff 最近兩日 →
  🟢新增成分 / 🔴剔除 / ⬆️⬇️增減碼。
跨基金:🔥 多檔主動 ETF「同時新增/加碼同一股」= 法人共識(最強訊號)。
"""
import os
import io
import sys
import glob
from datetime import date, timedelta
import requests
import pandas as pd

OUTDIR = "data/etf_holdings"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

# 主動式 ETF 監看清單(台股權益型;經理人選股)
# house 類型:
#   fhtrust  = 復華, /api/assetsExcel/{etf}/{YYYYMMDD}
#   ezmoney  = 統一, /ETF/Transaction/PCFExcelNPOI?fundCode={fc}&date={ROC}
#   direct   = 固定 URL (token 帶在裡面), 適用 CTBC 等. ⚠️ token 可能過期需重刷
FUNDS = [
    {"code": "00991A", "name": "復華未來50",   "house": "fhtrust", "etf": "ETF23"},
    {"code": "00998A", "name": "復華金融股息", "house": "fhtrust", "etf": "ETF24"},
    {"code": "00981A", "name": "統一台股增長", "house": "ezmoney", "fc": "49YTW"},
    {"code": "00403A", "name": "統一升級50",   "house": "ezmoney", "fc": "63YTW"},
    # ⚠️ token 若過期 → 打開 https://www.ctbcinvestments.com.tw/Etf/00406A/Combination
    # F12 Network 找 DownloadETFHoldingWeight 抓新 URL 更新
    # 尾綴 024 移除, base64 token 應以 == 結尾
    {"code": "00406A", "name": "中信主動式ETF", "house": "direct",
     "url": "https://www.ctbcinvestments.com.tw/API/etf/DownloadETFHoldingWeight?token=bwWJTtJZjUg2CtlP%2FI%2BOPSEucozq0mi7b0iB1O6GpbLR9vuI5ZQqPCNcpXYgoQYLMTY0MzkyMTkyNjUyNjY4OA%3D%3D"},
]


def roc(d):
    return f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"


import time

def fetch(fund, d, retries=3):
    """回傳 Excel bytes 或 None. 每檔都會 print 結果 (成功/失敗)."""
    s = requests.Session()
    s.headers["User-Agent"] = UA
    # 復華 fhtrust 需要 Referer 才穩 (不加會 RemoteDisconnected)
    if fund["house"] == "fhtrust":
        s.headers["Referer"] = "https://www.fhtrust.com.tw/"
        s.headers["Accept"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*"
    last_err = None
    for attempt in range(retries):
        try:
            if fund["house"] == "fhtrust":
                url = f"https://www.fhtrust.com.tw/api/assetsExcel/{fund['etf']}/{d:%Y%m%d}"
                r = s.get(url, timeout=30)
            elif fund["house"] == "direct":
                # token 在 URL 內, 無日期參數 → server 回最新持股
                # 有些 endpoint 只接受 POST (中信 CTBC), GET 失敗 fallback POST
                r = s.get(fund["url"], timeout=30)
                if r.status_code == 405:
                    r = s.post(fund["url"], timeout=30)
            else:
                url = "https://www.ezmoney.com.tw/ETF/Transaction/PCFExcelNPOI"
                r = s.get(url, params={"fundCode": fund["fc"], "date": roc(d), "specificDate": "true"}, timeout=30)
            if r.status_code == 200 and len(r.content) > 500:
                print(f"  {fund['code']} ✓ 抓到 ({len(r.content)/1024:.1f}KB)")
                return r.content
            last_err = f"HTTP {r.status_code}, content {len(r.content)} bytes"
            if r.status_code == 500 and fund["house"] == "direct":
                last_err += " (token 可能過期)"
                break  # 500 不用 retry
            if 400 <= r.status_code < 500:
                break  # 4xx 不用 retry
        except Exception as e:
            last_err = str(e)[:80]
        if attempt < retries - 1:
            time.sleep(2 ** attempt)  # backoff 1s, 2s
    print(f"  {fund['code']} ✗ 抓取失敗: {last_err}")
    return None


def _parse_one_sheet(content, sheet_name, code, tag):
    """單一 sheet parse 邏輯. 失敗回 None (不 print, 由外層決定)."""
    raw = pd.read_excel(io.BytesIO(content), header=None, sheet_name=sheet_name)
    hdr = None
    for i in range(min(30, len(raw))):
        if raw.iloc[i].astype(str).str.contains("代號").any():
            hdr = i; break
    if hdr is None:
        return None, "找不到「代號」表頭"
    df = pd.read_excel(io.BytesIO(content), header=hdr, sheet_name=sheet_name)

    def col(keys, exclude=None):
        for c in df.columns:
            cs = str(c)
            if exclude and any(x in cs for x in exclude):
                continue
            if any(k in cs for k in keys):
                return c
        return None
    # 排除「期貨代號」以避免抓到期貨 sheet
    ci = col(["代號"], exclude=["期貨"])
    cs = col(["股數"], exclude=["口數"])
    cn = col(["名稱"], exclude=["期貨"])
    cw = col(["權重", "比重"])
    if not (ci and cs):
        return None, f"缺欄位 代號={ci} 股數={cs} (欄位: {list(df.columns)})"
    out = pd.DataFrame({
        "代號": df[ci].astype(str).str.replace(r"\.0$", "", regex=True).str.strip(),
        "名稱": df[cn].astype(str).str.strip() if cn else "",
        "股數": pd.to_numeric(df[cs].astype(str).str.replace(",", ""), errors="coerce"),
        "權重": pd.to_numeric(df[cw].astype(str).str.replace("%", ""), errors="coerce") if cw else None,
    })
    before = len(out)
    out = out[out["代號"].str.match(r"^\d{4}$")].dropna(subset=["股數"])
    if out.empty:
        return None, f"filter 後空 (原 {before} rows)"
    return out.set_index("代號"), None


def parse(content, code=None):
    """兩家格式通用. 支援多 sheet: 逐個 sheet 試, 找到含 4 位數代號的 sheet 為止."""
    tag = f"[{code}]" if code else ""
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        print(f"  {tag} parse: Excel 讀不了 ({str(e)[:60]})")
        return None
    errors = []
    for sn in xl.sheet_names:
        try:
            df, err = _parse_one_sheet(content, sn, code, tag)
            if df is not None:
                if len(xl.sheet_names) > 1:
                    print(f"  {tag} parse: 使用 sheet「{sn}」({len(df)} 檔)")
                return df
            errors.append(f"sheet「{sn}」→ {err}")
        except Exception as e:
            errors.append(f"sheet「{sn}」→ 例外 {str(e)[:50]}")
    print(f"  {tag} parse 失敗 (試了 {len(xl.sheet_names)} 個 sheet): {'; '.join(errors)}")
    return None


def snapshot_path(code, d):
    return os.path.join(OUTDIR, f"{code}_{d:%Y%m%d}.csv")


def save_snapshot(code, d, df):
    os.makedirs(OUTDIR, exist_ok=True)
    df.reset_index().to_csv(snapshot_path(code, d), index=False, encoding="utf-8-sig")


def prev_snapshot(code, before):
    """讀此 code 的前一個快照(早於 before 的最新一個)。"""
    files = sorted(glob.glob(os.path.join(OUTDIR, f"{code}_*.csv")))
    files = [f for f in files if f < snapshot_path(code, before)]
    if not files:
        return None
    df = pd.read_csv(files[-1], dtype={"代號": str})
    return df.set_index("代號"), os.path.basename(files[-1])


def diff(cur, prev):
    sa, sb = set(prev.index), set(cur.index)
    new = [(s, cur.loc[s, "名稱"], cur.loc[s, "權重"]) for s in sb - sa]
    drop = [(s, prev.loc[s, "名稱"], prev.loc[s, "權重"]) for s in sa - sb]
    chg = []
    for s in sa & sb:
        p0 = prev.loc[s, "股數"]
        d = cur.loc[s, "股數"] - p0
        if not d or not p0:
            continue
        w0 = prev.loc[s, "權重"]
        # 期初基數過小(權重<0.1% 或 股數極微)→ 視為「建倉」,不算百分比
        # (否則 分母趨0 → +84344% 之類的爆表假訊號)
        build = (pd.notna(w0) and w0 < 0.1) or p0 < 1000
        pct = None if build else d / p0 * 100
        chg.append((s, cur.loc[s, "名稱"], pct, cur.loc[s, "權重"], build))
    # 建倉(pct=None)排最前,其餘按變動幅度大→小
    chg.sort(key=lambda x: (1e9 if x[2] is None else abs(x[2])), reverse=True)
    return new, drop, chg


def pull_day(d):
    """抓+存當日所有基金快照,回傳 {code: df}。"""
    out = {}
    for fund in FUNDS:
        content = fetch(fund, d)
        if content is None:
            continue
        cur = parse(content, code=fund["code"])
        if cur is None or cur.empty:
            print(f"  {fund['code']} parse 失敗, 跳過 (但 fetch 有拿到 bytes)")
            continue
        save_snapshot(fund["code"], d, cur)
        out[fund["code"]] = cur
    return out


def cross_fund(day_dfs):
    """跨基金分析:共同持有(共識核心)+ 買賣共識/分歧。"""
    codes = [c for c in day_dfs if not day_dfs[c].empty]
    if len(codes) < 2:
        return
    nm = {f["code"]: f["name"] for f in FUNDS}
    sets = [set(day_dfs[c].index) for c in codes]
    common = set.intersection(*sets)
    ref = day_dfs[codes[0]]
    print("🔥 跨基金『共同持有』(法人共識核心):")
    for s in sorted(common, key=lambda s: -(ref.loc[s, "權重"] if s in ref.index else 0))[:12]:
        ws = " | ".join(f"{nm[c][:2]}{day_dfs[c].loc[s, '權重']:.1f}%" for c in codes if s in day_dfs[c].index)
        nmstr = ref.loc[s, "名稱"] if s in ref.index else s
        print(f"   {s} {str(nmstr)[:5]:5s}  {ws}")
    print(f"   共同 {len(common)}檔 / 共 {len(codes)} 檔基金\n")


def diff_consensus(today):
    """以已存快照,算今日 vs 前一快照的跨基金買賣共識/分歧。"""
    buy, sell = {}, {}
    nm = {f["code"]: f["name"] for f in FUNDS}
    for fund in FUNDS:
        cur_f = snapshot_path(fund["code"], today)
        if not os.path.exists(cur_f):
            continue
        cur = pd.read_csv(cur_f, dtype={"代號": str}).set_index("代號")
        p = prev_snapshot(fund["code"], today)
        if not p:
            continue
        prev, pname = p
        new, drop, chg = diff(cur, prev)
        print(f"[{fund['code']} {fund['name']}] vs {pname}:")
        for s, n, w in new:
            print(f"   🟢新增 {s} {n}"); buy.setdefault(s, []).append(fund["name"])
        for s, n, w in drop:
            print(f"   🔴剔除 {s} {n}"); sell.setdefault(s, []).append(fund["name"])
        for s, n, pct, w, build in chg[:5]:
            if build:
                print(f"   🆕建倉 {s} {n} (現權重{w}%)")
                buy.setdefault(s, []).append(fund["name"])
            else:
                print(f"   {'⬆️加' if pct > 0 else '⬇️減'} {s} {n} {pct:+.0f}%")
                (buy if pct > 0 else sell).setdefault(s, []).append(fund["name"])
        if not (new or drop or chg):
            print("   (無變化)")
        print()
    # 共識(≥2檔同方向)+ 分歧(一買一賣)
    mb = {s: set(f) for s, f in buy.items() if len(set(f)) >= 2}
    ms = {s: set(f) for s, f in sell.items() if len(set(f)) >= 2}
    div = {s: (set(buy.get(s, [])), set(sell.get(s, []))) for s in set(buy) & set(sell)}
    if mb:
        print("🔥🔥 共識加碼(多檔同買):", "、".join(f"{s}({'/'.join(v)})" for s, v in mb.items()))
    if ms:
        print("🔥🔥 共識減碼(多檔同砍):", "、".join(f"{s}({'/'.join(v)})" for s, v in ms.items()))
    if div:
        print("⚖️ 分歧(有人買有人砍):", "、".join(f"{s}(買{','.join(b)}|砍{','.join(se)})" for s, (b, se) in div.items()))
    if not (mb or ms or div):
        print("(無跨基金共識/分歧訊號)")


def week_report():
    """讀所有已存快照,做近一週每檔淨變化 + 期初期末共識。"""
    nm = {f["code"]: f["name"] for f in FUNDS}
    print("\n========== 近期持股趨勢(全部已存快照)==========")
    for fund in FUNDS:
        files = sorted(glob.glob(os.path.join(OUTDIR, f"{fund['code']}_*.csv")))
        if len(files) < 2:
            continue
        a = pd.read_csv(files[0], dtype={"代號": str}).set_index("代號")
        b = pd.read_csv(files[-1], dtype={"代號": str}).set_index("代號")
        d0, d1 = files[0][-12:-4], files[-1][-12:-4]
        print(f"\n━━ {fund['code']} {fund['name']}  {d0}→{d1} ━━")
        new, drop, chg = diff(b, a)
        if new: print("  🟢期間新進:", "、".join(f"{s}{n}" for s, n, w in new))
        if drop: print("  🔴期間剔除:", "、".join(f"{s}{n}" for s, n, w in drop))
        for s, n, pct, w, build in [x for x in chg if x[2] is None or abs(x[2]) >= 10][:8]:
            if build:
                print(f"   🆕建倉 {s}{n} (現權重{w}%)")
            else:
                print(f"   {'⬆️加' if pct > 0 else '⬇️減'} {s}{n} {pct:+.0f}% (現權重{w}%)")


def _collect_diff(today):
    """收集每檔基金的 diff dict, 供 email/共識分析用
    return: (per_fund [{code, name, prev_date, new, drop, chg}], buy_map, sell_map)
    """
    per_fund = []
    buy, sell = {}, {}
    for fund in FUNDS:
        cur_f = snapshot_path(fund["code"], today)
        if not os.path.exists(cur_f):
            continue
        cur = pd.read_csv(cur_f, dtype={"代號": str}).set_index("代號")
        p = prev_snapshot(fund["code"], today)
        if not p:
            continue
        prev, pname = p
        new, drop, chg = diff(cur, prev)
        per_fund.append({"code": fund["code"], "name": fund["name"],
                         "prev": pname, "new": new, "drop": drop, "chg": chg})
        for s, n, w in new:
            buy.setdefault(s, {"name": n, "funds": []})["funds"].append(fund["name"])
        for s, n, w in drop:
            sell.setdefault(s, {"name": n, "funds": []})["funds"].append(fund["name"])
        for s, n, pct, w, build in chg[:5]:
            if build:
                buy.setdefault(s, {"name": n, "funds": []})["funds"].append(fund["name"])
            elif pct and pct > 20:
                buy.setdefault(s, {"name": n, "funds": []})["funds"].append(fund["name"])
            elif pct and pct < -20:
                sell.setdefault(s, {"name": n, "funds": []})["funds"].append(fund["name"])
    return per_fund, buy, sell


def _write_email_diff(today, pull_result=None):
    """產出 /tmp/etf_diff_subject.txt + _body.html
    只有結構性變化 (新進/剔除/共識) 才寫 subject → workflow 判斷是否寄信
    pull_result: {code: bool} 每檔今日抓取結果 (True=成功, False=失敗)
    """
    per_fund, buy, sell = _collect_diff(today)
    if not per_fund:
        print("⚠️ 無 diff 資料 (可能是第一次跑, 沒 prev), 不產 email")
        return

    # 共識 (≥2 檔同方向)
    consensus_buy = {s: v for s, v in buy.items() if len(set(v["funds"])) >= 2}
    consensus_sell = {s: v for s, v in sell.items() if len(set(v["funds"])) >= 2}
    divergent = {s: (buy[s], sell[s]) for s in set(buy) & set(sell)}

    # 判斷是否值得寄信: 有共識/新進/剔除/建倉/大幅變動 (>= 50%)
    total_new = sum(len(f["new"]) for f in per_fund)
    total_drop = sum(len(f["drop"]) for f in per_fund)
    total_build = sum(1 for f in per_fund for c in f["chg"] if c[4])  # build flag
    total_big = sum(1 for f in per_fund for c in f["chg"]
                    if c[2] is not None and abs(c[2]) >= 50)
    worth_send = bool(consensus_buy or consensus_sell or divergent
                      or total_new or total_drop or total_build or total_big)
    if not worth_send:
        print("⚠️ 今日僅小幅權重變化 (< 50%), 無結構性訊號, 不產 email")
        return

    parts = []
    if consensus_buy: parts.append(f"{len(consensus_buy)} 共識加碼")
    if consensus_sell: parts.append(f"{len(consensus_sell)} 共識減碼")
    if total_new: parts.append(f"{total_new} 新進")
    if total_drop: parts.append(f"{total_drop} 剔除")
    if total_build: parts.append(f"{total_build} 建倉")
    if total_big: parts.append(f"{total_big} 大幅動")
    subject = f"📊 主動 ETF 持股 diff — {' + '.join(parts)} ({today:%m/%d})"
    with open("/tmp/etf_diff_subject.txt", "w", encoding="utf-8") as f:
        f.write(subject)

    html = f"""<html><body style='font-family:-apple-system,sans-serif;max-width:900px'>
<h2>📊 主動式 ETF 持股 diff ({today:%Y-%m-%d})</h2>
<p style='color:#666'>對比昨日, 監看清單共 {len(FUNDS)} 檔, 今日 {len(per_fund)} 檔有可比 diff</p>
"""

    # 抓取狀態摘要 (讓失敗的檔透明)
    if pull_result is not None:
        html += "<h3>🔍 今日抓取狀態</h3>\n<table style='border-collapse:collapse'>\n"
        for fund in FUNDS:
            ok = pull_result.get(fund["code"], False)
            emoji = "✅" if ok else "❌"
            note = "" if ok else " <span style='color:#c33'>(見 console log)</span>"
            html += f"<tr><td style='padding:2px 8px'>{emoji}</td><td style='padding:2px 8px'>{fund['code']}</td><td style='padding:2px 8px'>{fund['name']}</td><td style='padding:2px 8px'>{note}</td></tr>\n"
        html += "</table>\n"

    if consensus_buy or consensus_sell or divergent:
        html += "<h3>🔥🔥 跨基金共識訊號 (≥2 檔同向)</h3>\n"
        if consensus_buy:
            html += "<div style='background:#e8f7e8;padding:10px;border-left:4px solid #2a2;margin:8px 0'>"
            html += "<b>🟢 共識加碼</b><ul>\n"
            for s, v in consensus_buy.items():
                html += f"<li><b>{s} {v['name']}</b> — {'/'.join(set(v['funds']))}</li>\n"
            html += "</ul></div>\n"
        if consensus_sell:
            html += "<div style='background:#fee;padding:10px;border-left:4px solid #c33;margin:8px 0'>"
            html += "<b>🔴 共識減碼</b><ul>\n"
            for s, v in consensus_sell.items():
                html += f"<li><b>{s} {v['name']}</b> — {'/'.join(set(v['funds']))}</li>\n"
            html += "</ul></div>\n"
        if divergent:
            html += "<div style='background:#fffbe0;padding:10px;border-left:4px solid #fa0;margin:8px 0'>"
            html += "<b>⚖️ 分歧 (有人買有人砍)</b><ul>\n"
            for s, (b_info, s_info) in divergent.items():
                nm = b_info["name"] or s_info["name"]
                html += f"<li><b>{s} {nm}</b> — 買:{'/'.join(set(b_info['funds']))} | 砍:{'/'.join(set(s_info['funds']))}</li>\n"
            html += "</ul></div>\n"

    html += "<h3>📋 各基金個別變化</h3>\n"
    for f in per_fund:
        rows = []
        for s, n, w in f["new"]:
            rows.append(f"<tr><td>🟢 新進</td><td>{s}</td><td>{n}</td><td style='text-align:right'>{w}%</td></tr>")
        for s, n, w in f["drop"]:
            rows.append(f"<tr><td>🔴 剔除</td><td>{s}</td><td>{n}</td><td style='text-align:right'>(原 {w}%)</td></tr>")
        for s, n, pct, w, build in f["chg"][:8]:
            if build:
                rows.append(f"<tr><td>🆕 建倉</td><td>{s}</td><td>{n}</td><td style='text-align:right'>權重 {w}%</td></tr>")
            elif pct and abs(pct) >= 10:
                arrow = "⬆️加" if pct > 0 else "⬇️減"
                rows.append(f"<tr><td>{arrow}</td><td>{s}</td><td>{n}</td><td style='text-align:right'>{pct:+.0f}% (權 {w}%)</td></tr>")
        if not rows:
            html += f"<h4>{f['code']} {f['name']}</h4><p style='color:#999'>(無顯著變化)</p>"
            continue
        html += f"<h4>{f['code']} {f['name']} <span style='color:#999;font-size:12px'>vs {f['prev']}</span></h4>\n"
        html += "<table style='border-collapse:collapse;width:100%'><tr style='background:#f0f0f0'>"
        html += "<th style='text-align:left;padding:4px'>動作</th><th style='text-align:left;padding:4px'>代號</th>"
        html += "<th style='text-align:left;padding:4px'>名稱</th><th style='text-align:right;padding:4px'>幅度</th></tr>\n"
        html += "\n".join(rows)
        html += "</table>\n"

    html += """
<div style='background:#f5f5ff;padding:12px;border-left:4px solid #55b;font-size:13px;margin-top:16px'>
<h4>📖 判讀說明</h4>
<p><b>🔥🔥 共識訊號</b>: 2 檔以上基金同時買/砍同檔 → 法人共識 (最值得跟)</p>
<p><b>⚖️ 分歧</b>: 一買一砍 → 各家看法不同, 可深入研究</p>
<p><b>🆕 建倉</b>: 新加入或原本權重 &lt; 0.1% → 經理人首次投入</p>
<p><b>⬆️加/⬇️減</b>: 顯示 &gt;= 10% 的變化, 過濾雜訊</p>
</div>
</body></html>"""

    with open("/tmp/etf_diff_body.html", "w", encoding="utf-8") as f:
        f.write(html)
    print(f"→ /tmp/etf_diff_subject.txt + _body.html ({subject})")


def main():
    backfill = int(os.environ.get("ETF_BACKFILL_DAYS", "0") or "0")
    if backfill > 0:
        # 回補:從 today 往回抓 backfill 天(假日自動跳過),建立一週快照
        end = date.today()
        print(f"=== 回補近 {backfill} 天主動式ETF快照 ===\n")
        for i in range(backfill, -1, -1):
            d = end - timedelta(days=i)
            got = pull_day(d)
            print(f"  {d:%Y/%m/%d}: {len(got)} 檔基金有資料")
        week_report()
        return

    _d = os.environ.get("ETF_DATE", "").strip()
    today = pd.to_datetime(_d, format="%Y%m%d").date() if _d else date.today()
    print(f"=== 主動式 ETF 持股監看 {today:%Y/%m/%d} ===\n")
    day_dfs = pull_day(today)
    if not day_dfs:
        print("今日無資料(假日/未公告)"); return
    diff_consensus(today)
    print()
    cross_fund(day_dfs)
    print()
    # pull_result 傳給 email 顯示每檔抓取狀態
    pull_result = {f["code"]: (f["code"] in day_dfs) for f in FUNDS}
    _write_email_diff(today, pull_result=pull_result)


if __name__ == "__main__":
    main()
