# -*- coding: utf-8 -*-
"""
體檢系統 回歸測試 (Health Check Regression Test)
=====================================================================
拿 regression_cases.yaml 的 11 檔金樣本,比對 data/台股_體檢總表.xlsx 的當前判讀。
改完 code → 重跑 stock_health_check.py → 跑這支,確認系統沒退化。

期望語法(每檔可選用):
  評等_在: [A, B]      評等必須落在清單內
  扣分_至多: -10       ⑪動態惡化扣分 <= 此值(更負=扣更多也OK)
  漏洞_含: "償債"      主要漏洞字串需含此關鍵字
  鬧鐘_含: "陷阱"      鬧鐘字串需含此關鍵字
  鬧鐘_不含: "陷阱"    鬧鐘字串不可含此關鍵字(可加 _不含2 _不含3)

退出碼:全過=0,有 FAIL=1(可接 CI gate)
"""
import sys
import os
import pandas as pd
import yaml

SRC = "data/台股_體檢總表.xlsx"
CASES = "regression_cases.yaml"
SNAP = "snapshot_cases.yaml"


def load():
    df = pd.read_excel(SRC, "體檢總表")
    df["代號"] = df["代號"].astype(str)
    with open(CASES, encoding="utf-8") as f:
        cases = yaml.safe_load(f)["cases"]
    return df, cases


def check_one(row, exp):
    """回傳 (通過bool, 失敗訊息list)。row=該檔體檢資料(Series),exp=期望dict。"""
    fails = []
    grade = str(row.get("評等", ""))
    pen = pd.to_numeric(row.get("⑪動態惡化扣分"), errors="coerce")
    leak = str(row.get("主要漏洞", "") or "")
    alarm = str(row.get("鬧鐘", "") or "")

    if "評等_在" in exp and grade not in exp["評等_在"]:
        fails.append(f"評等={grade} 不在 {exp['評等_在']}")
    if "扣分_至多" in exp:
        thr = exp["扣分_至多"]
        if pd.isna(pen) or pen > thr:
            fails.append(f"⑪扣分={pen}(需 ≤{thr})")
    if "漏洞_含" in exp and exp["漏洞_含"] not in leak:
        fails.append(f"主要漏洞缺『{exp['漏洞_含']}』(實際:{leak[:30]})")
    if "鬧鐘_含" in exp and exp["鬧鐘_含"] not in alarm:
        fails.append(f"鬧鐘缺『{exp['鬧鐘_含']}』(實際:{alarm})")
    for k in ("鬧鐘_不含", "鬧鐘_不含2", "鬧鐘_不含3"):
        if k in exp and exp[k] in alarm:
            fails.append(f"鬧鐘不該含『{exp[k]}』(實際:{alarm})")
    return (len(fails) == 0), fails


def check_snapshot(row, frozen):
    """快照漂移檢查:評等/鬧鐘/漏洞關鍵字 是否跟凍結值一致。回傳 (一致bool, 漂移list)。"""
    drift = []
    grade = str(row.get("評等", ""))
    alarm = (str(row.get("鬧鐘")) if pd.notna(row.get("鬧鐘")) else None)
    leak = str(row.get("主要漏洞", "") or "")
    if "評等" in frozen and grade != frozen["評等"]:
        drift.append(f"評等 {frozen['評等']}→{grade}")
    if "鬧鐘" in frozen and alarm != frozen["鬧鐘"]:
        drift.append(f"鬧鐘 {frozen['鬧鐘']}→{alarm}")
    for kw in frozen.get("漏洞關鍵字", []):
        if kw not in leak:
            drift.append(f"漏洞少了『{kw}』")
    return (len(drift) == 0), drift


def main():
    df, cases = load()
    npass = nfail = 0
    print(f"=== ① 正確性金樣本:{len(cases)} 檔(人工驗證,守對錯)===\n")
    for c in cases:
        sid = c["sid"]
        hit = df[df["代號"] == sid]
        if hit.empty:
            print(f"  ⚠️ {sid} {c['name']:6s} 體檢表查無此檔(資料缺失?)")
            nfail += 1
            continue
        ok, fails = check_one(hit.iloc[0], c.get("期望", {}))
        if ok:
            print(f"  ✅ {sid} {c['name']:6s} [{c.get('類別','')}]")
            npass += 1
        else:
            print(f"  ❌ {sid} {c['name']:6s} [{c.get('類別','')}]")
            for f in fails:
                print(f"        └ {f}")
            nfail += 1

    # ② 快照守門(非人工驗證,只守「判讀沒漂移」;漂移=警告不算 FAIL)
    ndrift = 0
    if os.path.exists(SNAP):
        with open(SNAP, encoding="utf-8") as f:
            snaps = (yaml.safe_load(f) or {}).get("snapshots", [])
        print(f"\n=== ② 快照守門:{len(snaps)} 檔(偵錯批,守判讀漂移)===\n")
        for s in snaps:
            hit = df[df["代號"] == s["sid"]]
            if hit.empty:
                continue
            ok, drift = check_snapshot(hit.iloc[0], s.get("凍結", {}))
            if not ok:
                print(f"  🔄 {s['sid']} {s['name']:6s} 判讀漂移:")
                for d in drift:
                    print(f"        └ {d}")
                ndrift += 1
        if not ndrift:
            print(f"  ✅ {len(snaps)} 檔判讀無漂移")

    print(f"\n=== 結果:正確性 {npass} 過 / {nfail} 失敗;快照漂移 {ndrift} 檔 ===")
    if nfail:
        print("❌ 正確性退化!改 code 改壞了已驗證判讀,必修。")
    if ndrift:
        print("🔄 快照漂移(非必錯):複查是否為預期改善,確認後跑 make_snapshot.py 更新基準。")
    sys.exit(1 if nfail else 0)   # 只有正確性 FAIL 才讓 CI 紅;漂移僅警告


if __name__ == "__main__":
    main()
