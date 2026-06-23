# -*- coding: utf-8 -*-
"""
一次性回填:市值(億)補進既有快取,供 0050 納入預測用。
市值 = (股本 ÷ 10) × 收盤;股本(CommonStocks)在資產負債表,每檔 1 次 API。
已有「市值(億)」者跳過。跑完重建一次 Excel。
"""
import time
import pandas as pd
import fetch_fundamentals_tw as F

MAX_RUNTIME_MIN = 28


def latest_capital(bal_df):
    piv = F.pivot(bal_df)
    cap = F.pick(piv, "CommonStocks", "CommonStock", "CapitalStock", "Capital", "ShareCapital")
    cap = pd.to_numeric(cap, errors="coerce").dropna()
    if not len(cap):
        return None
    return float(cap.iloc[-1])


def main():
    t0 = time.time()
    dl = F.make_loader()
    namemap = F.load_names(dl)
    todo = []
    for sid in F.PICKS:
        c = F.load_cache(sid)
        if c is None:
            continue
        row = c.get("row") or {}
        if row.get("市值(億)") is not None:
            continue
        if row.get("收盤") is None:
            continue   # 沒收盤就算不出市值,下次完整跑再補
        todo.append(sid)
    print(f"待回填市值 {len(todo)} 檔")

    for i, sid in enumerate(todo, 1):
        if time.time() - t0 > MAX_RUNTIME_MIN * 60:
            print(f"⏲ 達 {MAX_RUNTIME_MIN} 分上限,剩 {len(todo)-i+1} 檔下輪補"); break
        tries = 0
        while True:
            try:
                bal = dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=F.START_DATE)
                cap = latest_capital(bal)
                c = F.load_cache(sid) or {}
                row = c.get("row") or {}
                close = row.get("收盤")
                if cap and close:
                    shares = cap / 10.0
                    mcap = round(shares * float(close) / 1e8, 1)
                    row["市值(億)"] = mcap
                    c["row"] = row
                    F.save_cache(sid, c)
                    print(f"[{i}/{len(todo)}] {sid} {namemap.get(sid,sid)} 市值 {mcap} 億")
                else:
                    print(f"[{i}/{len(todo)}] {sid} 抓不到股本/收盤,跳過")
                break
            except Exception as e:
                if F._is_rate_limit(e) and tries < 2:
                    tries += 1; print(f"  ⏸ 額度,睡90s({tries}/2)"); time.sleep(90); continue
                if F._is_rate_limit(e):
                    print(f"  ↷ {sid} 額度未恢復,跳過"); break
                print(f"  ! {sid} 失敗:{e}"); break
        time.sleep(F.RATE_SLEEP)

    n = F.build_output(namemap)
    print(f"完成,重建 {n} 檔")


if __name__ == "__main__":
    main()
