# -*- coding: utf-8 -*-
"""
一次性回填:逐季毛利率(q_gm)補進既有快取,供拐點掃描算「毛利拐頭」。
只抓損益表(每檔 1 次 API),不動既有 row/hist/年度資料 → 220 檔一輪即可、不破壞現有輸出。
跑完重建一次 Excel(含『逐季毛利率』分頁)。可重複執行,已補的會跳過。
"""
import time
import pandas as pd
import fetch_fundamentals_tw as F   # 重用 loader/pivot/cache/PICKS/設定

MAX_RUNTIME_MIN = 28


def main():
    t0 = time.time()
    dl = F.make_loader()
    namemap = F.load_names(dl)
    todo = []
    for sid in F.PICKS:
        c = F.load_cache(sid)
        if c is not None and not c.get("q_gm"):   # 有快取但沒逐季毛利 → 待補
            todo.append(sid)
    print(f"待回填逐季毛利 {len(todo)} 檔")

    for i, sid in enumerate(todo, 1):
        if time.time() - t0 > MAX_RUNTIME_MIN * 60:
            print(f"⏲ 達 {MAX_RUNTIME_MIN} 分上限,先收尾(剩 {len(todo)-i+1} 檔下輪補)"); break
        tries = 0
        while True:
            try:
                inc = dl.taiwan_stock_financial_statement(stock_id=sid, start_date=F.START_DATE)
                piv = F.pivot(inc)
                rev = F.pick(piv, "Revenue")
                gp  = F.pick(piv, "GrossProfit")
                gm = (gp / rev * 100).dropna().tail(8)
                q3 = {str(d)[:10]: round(float(x), 2) for d, x in gm.items()}
                c = F.load_cache(sid) or {}
                c["q_gm"] = q3
                F.save_cache(sid, c)
                print(f"[{i}/{len(todo)}] {sid} {namemap.get(sid,sid)} 季毛利 {len(q3)} 季")
                break
            except Exception as e:
                if F._is_rate_limit(e) and tries < 2:
                    tries += 1; print(f"  ⏸ 額度,睡90s短重試({tries}/2)"); time.sleep(90); continue
                if F._is_rate_limit(e):
                    print(f"  ↷ {sid} 額度未恢復,跳過(下輪補)"); break
                print(f"  ! {sid} 失敗:{e}"); break
        time.sleep(F.RATE_SLEEP)

    n = F.build_output(namemap)        # 重建 Excel(含逐季毛利率分頁)
    print(f"完成,重建 {n} 檔")


if __name__ == "__main__":
    main()
