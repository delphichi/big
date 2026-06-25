# -*- coding: utf-8 -*-
"""
產生 體檢快照 (Snapshot Generator)
=====================================================================
把指定一批股票「目前的體檢判讀」凍結成 snapshot_cases.yaml。
用途:這些股票尚未人工驗證對錯,但凍結現況後,日後任何 code 改動
      若讓它們的「評等/鬧鐘/主要漏洞」漂移,test_health_regression.py
      會標記出來給你複查 —— 這叫「快照守門 / characterization test」。

與 regression_cases.yaml(11檔人工驗證的正確性金樣本)區別:
  - regression_cases  = 正確性 oracle(我們確定八方該綠、鈊象該抓雷)
  - snapshot_cases    = 現況快照(只保證「沒變」,不保證「對」)

用法:python make_snapshot.py        # 重新凍結 SNAP_SIDS 的現況
"""
import pandas as pd
import yaml

SRC = "data/台股_體檢總表.xlsx"
OUT = "snapshot_cases.yaml"

# 偵錯批 36 檔(營建/金融/電子/半導體設備材料)
SNAP_SIDS = [
    "2515","3703","6177","2509","1438","5213","2882","2881","2891","2886",
    "5880","2880","4938","2356","2301","5309","6488","5292","8176","6944",
    "7818","2013","6486","1521","6196","6826","7703","4755","4722","2422",
    "3680","1560","1764","3532","6640","6937",
]


def main():
    df = pd.read_excel(SRC, "體檢總表")
    df["代號"] = df["代號"].astype(str)
    snaps = []
    miss = []
    for sid in SNAP_SIDS:
        hit = df[df["代號"] == sid]
        if hit.empty:
            miss.append(sid)
            continue
        r = hit.iloc[0]
        snap = {
            "sid": sid,
            "name": str(r.get("名稱", "")),
            "凍結": {
                "評等": str(r.get("評等", "")),
                "鬧鐘": (str(r.get("鬧鐘")) if pd.notna(r.get("鬧鐘")) else None),
            },
        }
        # 主要漏洞:存「關鍵字」而非整串(整串太脆,數字微動就誤報)
        leak = str(r.get("主要漏洞", "") or "")
        keys = [k for k in ("EPS連年衰退", "毛利壓歷史低檔", "含金量", "高槓桿",
                            "短期償債", "存貨爆衝", "月營收轉負", "ROE滑落",
                            "營收5年萎縮", "EPS資料不足") if k in leak]
        if keys:
            snap["凍結"]["漏洞關鍵字"] = keys
        snaps.append(snap)

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("# 體檢快照 — 偵錯批現況凍結(非人工驗證,只守「判讀沒漂移」)\n")
        f.write("# 由 make_snapshot.py 產生。判讀有意改變時,重跑此腳本更新基準。\n")
        yaml.safe_dump({"snapshots": snaps}, f, allow_unicode=True, sort_keys=False)

    print(f"已凍結 {len(snaps)} 檔 → {OUT}")
    if miss:
        print(f"⚠️ {len(miss)} 檔尚未在體檢表(需先加PICKS重抓):{miss}")


if __name__ == "__main__":
    main()
