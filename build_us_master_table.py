# -*- coding: utf-8 -*-
"""
美股全市場總表 build_us_master_table.py
=======================================================================
不再只輸出「通過初篩」的,而是把全 ~10400 檔 SEC 快取**全部攤開**,每檔標清楚
狀態與原因,讓人自己看/篩,沒有東西被規則偷偷丟掉(「全盤」原則)。

分類:
  ✅通過            ROE≥15或ROIC≥12 + 毛利/含金量都OK且有值
  🟡夠力但資料不全   ROE/ROIC夠力, 但毛利或含金量在SEC算不出(None)→ 留待FMP深查
  🟠明確瑕疵         ROE/ROIC夠力, 但毛利明確≤0 或 含金量明確<0.8
  ⚪不夠力           ROE<15 且 ROIC<12
  🔴爆表失真         ROE>80 或 ROIC>60(微型股分母失真)
  ⚫算不出           SEC XBRL 標記不一致, 三大指標抓不到
  ▫️非普通股         權證/單位/特別股(代號含-或過長)

輸出:data/美股_全市場總表.xlsx(可排序篩選)+ 各類統計
"""
import os
import json
import pandas as pd

from us_revenue_yoy_scanner import load_cik_map

CACHE_DIR = "data/_profit_cache_us"
OUT = "data/美股_全市場總表.xlsx"
ROE_MIN, ROIC_MIN, CASH_MIN = 15.0, 12.0, 0.8


def load_watch():
    have = set()
    for f in ("tickers_us_core.txt", "tickers_us.txt"):
        if os.path.exists(f):
            for line in open(f, encoding="utf-8"):
                s = line.split("#")[0].strip().upper()
                if s:
                    have.add(s)
    return have


def classify(c, t):
    if ("-" in t) or (len(t) > 5) or (not t.isalpha()):
        return "▫️非普通股"
    if not c or c.get("error") or c.get("skip"):
        return "⚫算不出"
    roe, roic = c.get("roe_pct"), c.get("roic_pct")
    cash, gm = c.get("ocf_to_ni"), c.get("gross_margin")
    if (roe is not None and roe > 80) or (roic is not None and roic > 60):
        return "🔴爆表失真"
    # 資料矛盾:負ROE卻正高ROIC = SEC符號/規模錯, 或負淨值(庫藏股)→ 不可當好公司, 須FMP驗
    if (roe is not None and roe < 0) and (roic is not None and roic >= ROIC_MIN):
        return "⚠️資料矛盾(負ROE)"
    roe_ok = roe is not None and roe >= ROE_MIN
    roic_ok = roic is not None and roic >= ROIC_MIN
    if not (roe_ok or roic_ok):
        return "⚪不夠力"
    if (gm is not None and gm <= 0) or (cash is not None and cash < CASH_MIN):
        return "🟠明確瑕疵"
    if gm is None or cash is None:
        return "🟡夠力但資料不全"
    return "✅通過"


def main():
    cikmap = load_cik_map()
    names = {}
    # company_tickers.json: ticker→{cik,title};load_cik_map 可能只回 ticker→cik,名稱另查
    try:
        names = {t: (v.get("title") if isinstance(v, dict) else "") for t, v in cikmap.items()}
    except Exception:
        names = {}
    watch = load_watch()

    rows = []
    for t in sorted(cikmap.keys()):
        p = os.path.join(CACHE_DIR, f"{t}.json")
        c = None
        if os.path.exists(p):
            try:
                c = json.load(open(p, encoding="utf-8"))
            except Exception:
                c = None
        if c is None:
            continue  # 還沒抓到的略過(只列已有快取的)
        cls = classify(c, t)
        rows.append({
            "代號": t,
            "名稱": names.get(t, ""),
            "分類": cls,
            "ROE": c.get("roe_pct"),
            "ROIC": c.get("roic_pct"),
            "含金量": c.get("ocf_to_ni"),
            "毛利": c.get("gross_margin"),
            "FCF(B)": c.get("fcf_ttm_b"),
            "在觀察名單": "✔" if t in watch else "",
        })
    df = pd.DataFrame(rows)
    order = ["✅通過", "🟡夠力但資料不全", "⚠️資料矛盾(負ROE)", "🟠明確瑕疵", "🔴爆表失真", "⚪不夠力", "⚫算不出", "▫️非普通股"]
    df["_o"] = df["分類"].apply(lambda x: order.index(x) if x in order else 99)
    df = df.sort_values(["_o", "ROIC"], ascending=[True, False]).drop(columns="_o")

    stat = df["分類"].value_counts().reindex(order).fillna(0).astype(int)
    # 候選池 = 通過 + 夠力但資料不全(後者待 FMP 補)
    pool = df[df["分類"].isin(["✅通過", "🟡夠力但資料不全"])]
    pool_new = pool[pool["在觀察名單"] == ""]

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="全市場總表", index=False)
        pool.to_excel(xw, sheet_name="候選池(通過+待查)", index=False)
        pool_new.to_excel(xw, sheet_name="候選池_遺珠", index=False)
        stat.rename("檔數").reset_index().rename(columns={"index": "分類"}).to_excel(
            xw, sheet_name="分類統計", index=False)

    print(f"完成 → {OUT}(總 {len(df)} 檔)")
    for k in order:
        print(f"  {k}: {int(stat[k])}")
    print(f"\n候選池(通過+夠力待查)= {len(pool)} 檔 / 其中遺珠(不在觀察名單)= {len(pool_new)} 檔")
    print("候選池前 15(按 ROIC):")
    print(pool.head(15)[["代號", "名稱", "分類", "ROE", "ROIC", "含金量", "毛利"]].to_string(index=False))


if __name__ == "__main__":
    main()
