# -*- coding: utf-8 -*-
"""
台股 PE / PEG / Forward PE 監看表 tw_pe_monitor.py
=======================================================================
監看清單 97 檔(預設與 fetch_fundamentals_tw PICKS 同步),
抓即時 PER + 從體檢總表撈 ForwardPE / PEG / 估值鬧鐘,
輸出 data/台股PE監看表.xlsx,跟前次快照比漲跌。

跑法:
  python tw_pe_monitor.py
  WATCHLIST=2330,3017 python tw_pe_monitor.py   # 自訂清單
"""
import os
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

TOKEN = os.environ.get("FINMIND_TOKEN", "")
BASE = "https://api.finmindtrade.com/api/v4/data"
SRC = "data/台股_體檢總表.xlsx"
DST = "data/台股PE監看表.xlsx"
WORKERS = int(os.environ.get("MONITOR_WORKERS", "6"))

DEFAULT_WATCH = """
3017 3653 2383 6139 8210 8996 6442 2360 2345 6223 3324
2330 2308 6196 3583 6197 3044 2059 5434 1513 1560 5340 1519 3008
2640 3029 4506 3689 5519 2421 1618 1215 2618 5904 1232 6189 5478 4527
6788 2753 6515 3045 4129 2912 1514 3147 3402
6274 2472 3260 3406 3661 2327 2408 3711 3293 1503 2535 2476 6263 1773 1590
4303 5269 6446 4979 3450 3081 8271 2382 2379 3231 2301
5511 8926 1616 4933 3596 3227 6285 6510 2356 2317 5243 5225 3188 2305 8086
2395 2412 9917 6739 2645 6690 4569 4904 2344
""".split()


def get(dataset, **params):
    params["dataset"] = dataset
    if TOKEN: params["token"] = TOKEN
    for attempt in range(3):
        try:
            r = requests.get(BASE, params=params, timeout=15)
        except Exception:
            time.sleep(1); continue
        if r.status_code == 402: time.sleep(5); continue
        if r.status_code != 200: return None
        j = r.json()
        if j.get("status") != 200: return None
        return j.get("data", [])
    return None


def fetch_latest(code):
    """抓即時報價 + PER/PBR"""
    today = datetime.now(timezone.utc).date()
    start = (today - timedelta(days=10)).isoformat()
    try:
        # 報價
        p = get("TaiwanStockPrice", data_id=code, start_date=start)
        price = None
        if p:
            df = pd.DataFrame(p).sort_values("date")
            price = float(df.iloc[-1].get("close", 0))
        # PER / PBR
        per = get("TaiwanStockPER", data_id=code, start_date=start)
        pe = pbr = None
        if per:
            df = pd.DataFrame(per).sort_values("date").iloc[-1]
            pe = float(df.get("PER", 0)) or None
            pbr = float(df.get("PBR", 0)) or None
        return code, price, pe, pbr
    except Exception:
        return code, None, None, None


def alarm(pe, fwd_pe, peg):
    """估值鬧鐘"""
    if peg is not None and peg > 0:
        if peg < 1.0: return "🟢未來便宜"
        if peg < 1.5: return "🟢成長未反映"
        if peg < 2.0: return "🟡未來合理"
        if peg < 3.0: return "🟠未來偏貴"
        return "🔴未來過熱"
    if fwd_pe is not None and fwd_pe > 0:
        if fwd_pe < 12: return "🟢未來便宜"
        if fwd_pe < 18: return "🟡未來合理"
        if fwd_pe < 30: return "🟠未來偏貴"
        return "🔴未來過熱"
    if pe is not None and pe > 0:
        if pe < 12: return "🟢便宜"
        if pe < 20: return "🟡合理"
        if pe < 35: return "🟠偏貴"
        return "🔴過熱"
    return "—"


def main():
    if not TOKEN: print("⚠️ 未設 FINMIND_TOKEN(會被速率限制)")

    watch = os.environ.get("WATCHLIST", "").strip()
    if watch:
        codes = [s.strip() for s in watch.replace(",", " ").split() if s.strip()]
    else:
        codes = DEFAULT_WATCH
    codes = list(dict.fromkeys(codes))
    print(f"監看 {len(codes)} 檔(平行 {WORKERS})")

    # 從體檢總表撈 名稱 / 評等 / 品質 / EPS3y / ForwardPE / 估值 / 主要漏洞 / 殖利率 等
    base = pd.read_excel(SRC, sheet_name="體檢總表")
    base["代號"] = base["代號"].astype(str)
    keep = [c for c in ["代號","名稱","產業","評等","品質總分","EPS近3y%","ROE","含金量",
                        "營收5yCAGR","月營收YoY","PER","PBR","ForwardPE","PEG","估值",
                        "未來估值","殖利率%","循環股","主要漏洞"] if c in base.columns]
    base = base[base["代號"].isin(codes)][keep]

    # 平行抓即時報價
    quotes = {}
    print("抓即時報價...")
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fetch_latest, c): c for c in codes}
        done = 0
        for fut in as_completed(futs):
            code, price, pe, pbr = fut.result()
            quotes[code] = {"當前股價": price, "PER即時": pe, "PBR即時": pbr}
            done += 1
            if done % 30 == 0: print(f"  [{done}/{len(codes)}]")

    # 組合
    rows = []
    for c in codes:
        b = base[base["代號"] == c]
        b = b.iloc[0].to_dict() if len(b) else {"代號": c, "名稱": "(不在總表)"}
        q = quotes.get(c, {})
        b.update(q)
        # PEG 即時:用即時 PER / 過去 3y EPS%(粗算)
        peg_table = b.get("PEG")
        fwd = b.get("ForwardPE")
        if pd.notna(peg_table) and peg_table:
            b["PEG_使用"] = float(peg_table)
        else:
            b["PEG_使用"] = None
        b["估值鬧鐘"] = alarm(b.get("PER即時"), fwd, b.get("PEG_使用"))
        rows.append(b)

    df = pd.DataFrame(rows)

    # 與前次比對
    if os.path.exists(DST):
        try:
            prev = pd.read_excel(DST, sheet_name="監看表")
            prev["代號"] = prev["代號"].astype(str)
            prev = prev[["代號","當前股價"]].rename(columns={"當前股價":"前次股價"})
            df = df.merge(prev, on="代號", how="left")
            df["漲跌%"] = ((df["當前股價"] - df["前次股價"]) / df["前次股價"] * 100).round(2)
        except Exception:
            pass

    # 排序:過熱在前,綠燈在後
    order = {"🔴未來過熱":0,"🔴過熱":1,"🟠未來偏貴":2,"🟠偏貴":3,
             "🟡未來合理":4,"🟡合理":5,"🟢成長未反映":6,"🟢未來便宜":7,"🟢便宜":8,"—":9}
    df["_o"] = df["估值鬧鐘"].map(lambda x: order.get(x, 9))
    df = df.sort_values(["_o","品質總分"], ascending=[True, False]).drop(columns=["_o"])

    front = ["代號","名稱","產業","評等","品質總分","當前股價","PER即時","PBR即時",
             "ForwardPE","EPS近3y%","PEG_使用","估值鬧鐘"]
    if "漲跌%" in df.columns: front.append("漲跌%")
    if "前次股價" in df.columns: front.append("前次股價")
    front += ["殖利率%","循環股","主要漏洞"]
    front = [c for c in front if c in df.columns]
    rest = [c for c in df.columns if c not in front]
    df = df[front + rest]

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="監看表", index=False)
        for k, label in [("🟢未來便宜","買進_便宜"),("🟢成長未反映","買進_成長未反映"),
                         ("🔴未來過熱","警示_過熱"),("🟠未來偏貴","警示_偏貴")]:
            sub = df[df["估值鬧鐘"] == k]
            if len(sub): sub.to_excel(xw, sheet_name=label[:31], index=False)

    print(f"\n→ 已輸出 {DST}\n")
    print("估值鬧鐘分布:"); print(df["估值鬧鐘"].value_counts().to_string())
    buy = df[df["估值鬧鐘"].isin(["🟢成長未反映","🟢未來便宜","🟢便宜"])]
    if len(buy):
        print(f"\n🟢 買進信號 {len(buy)} 檔 TOP 15:")
        c = [x for x in ["代號","名稱","評等","品質總分","當前股價","PER即時","ForwardPE","PEG_使用","估值鬧鐘"] if x in buy.columns]
        print(buy[c].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
