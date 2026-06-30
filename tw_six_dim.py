# -*- coding: utf-8 -*-
"""
台股六維交叉 tw_six_dim.py
=======================================================================
合併 4 個 dashboard 輸出, 產出 100 檔六維評分:

  1. 成長 ← data/台股100檔_加速度分類.xlsx
  2. 籌碼 ← data/台股_籌碼儀表板.xlsx (三大法人/外資佔比/融資/借券)
  3. 報酬 ← data/台股_報酬雷達.xlsx (1y/3y 報酬 + TAIEX 超額)
  4. 警戒 ← data/台股_警戒掃描.xlsx (處置/借券爆量計數)
  5. 評等 ← 成長分類已 merge
  6. 品質 ← A 級且品質 90+ 加分

跑法: python tw_six_dim.py
輸出: data/台股100檔_六維交叉.xlsx
"""
import os
import pandas as pd

GROW_SRC = "data/台股100檔_加速度分類.xlsx"
CHIP_SRC = "data/台股_籌碼儀表板.xlsx"
RET_SRC = "data/台股_報酬雷達.xlsx"
ALERT_SRC = "data/台股_警戒掃描.xlsx"
DST = "data/台股100檔_六維交叉.xlsx"


def safe_read(path, sheet=None):
    if not os.path.exists(path): return pd.DataFrame()
    try:
        if sheet:
            try: return pd.read_excel(path, sheet_name=sheet)
            except Exception: return pd.read_excel(path)
        return pd.read_excel(path)
    except Exception as e:
        print(f"⚠️ {path}: {e}")
        return pd.DataFrame()


def six_score(r):
    """六維評分 + 訊號"""
    s = 0; flags = []

    # 1. 成長
    cat = str(r.get("分類",""))
    if "加速器" in cat: s += 2; flags.append("🚀加速")
    elif "高飛" in cat: s += 1; flags.append("✈️高飛")
    elif "失速" in cat: s -= 2; flags.append("💀失速")
    elif "減速" in cat: s -= 1; flags.append("📉減速")

    # 2. 籌碼分
    c = r.get("籌碼分")
    if pd.notna(c):
        c = int(c)
        if c >= 3: s += 2; flags.append("🟢籌碼強")
        elif c >= 1: s += 1
        elif c <= -2: s -= 2; flags.append("🔴籌碼弱")

    # 3. 報酬 1y 超額(對 TAIEX)
    e = r.get("1y超額%")
    if pd.notna(e):
        if e > 100: s += 2; flags.append("🏆暴贏大盤")
        elif e > 0: s += 1
        elif e < -50: s -= 2; flags.append("📉跌輸大盤")
        elif e < 0: s -= 1

    # 4. 警戒筆數
    a = int(r.get("警戒筆數", 0) or 0)
    if a >= 3: s -= 2; flags.append(f"⚠️警戒{a}次")
    elif a >= 1: s -= 1; flags.append(f"⚠️警戒{a}")

    # 5. 評等 A + 品質 90+ 加分
    if str(r.get("評等","")) == "A" and (r.get("品質", 0) or 0) >= 90:
        s += 1; flags.append("🏅A品90+")

    return s, " ".join(flags) if flags else "—"


def main():
    grow = safe_read(GROW_SRC)
    chip = safe_read(CHIP_SRC, sheet="總覽")
    ret  = safe_read(RET_SRC, sheet="總覽")
    alert= safe_read(ALERT_SRC, sheet="警戒總覽")

    if grow.empty:
        print("⚠️ 缺成長分類, 無法跑"); return

    # 統一 代號 str
    for d in [grow, chip, ret, alert]:
        if not d.empty and "代號" in d.columns:
            d["代號"] = d["代號"].astype(str)

    # 主表骨架
    base_cols = [c for c in ["代號","名稱","評等","品質","營收10y","營收5y","營收3y","淨利3y","分類"] if c in grow.columns]
    master = grow[base_cols].copy()
    print(f"成長分類: {len(master)} 檔")

    # merge 籌碼
    if not chip.empty:
        keep = [c for c in ["代號","外資90d淨","投信90d淨","三大90d淨","外資持股Δpp","融資Δ%","借券90d量","籌碼分","訊號"] if c in chip.columns]
        chip_sub = chip[keep].rename(columns={"訊號":"籌碼訊號"})
        master = master.merge(chip_sub, on="代號", how="left")
        print(f"籌碼: 已 merge")

    # merge 報酬
    if not ret.empty:
        keep = [c for c in ["代號","最新價","1y報酬%","3y報酬%","5y報酬%","1y年化%","3y年化%","5y年化%",
                            "1y超額%","3y超額%","5y超額%","5y最大回撤%"] if c in ret.columns]
        master = master.merge(ret[keep], on="代號", how="left")
        print(f"報酬: 已 merge")

    # 警戒筆數
    if not alert.empty and "代號" in alert.columns:
        ac = alert.groupby("代號").size().reset_index(name="警戒筆數")
        master = master.merge(ac, on="代號", how="left")
        master["警戒筆數"] = master["警戒筆數"].fillna(0).astype(int)
        print(f"警戒: 已 merge")
    else:
        master["警戒筆數"] = 0

    # 六維評分
    master["六維分"], master["六維訊號"] = zip(*master.apply(six_score, axis=1))
    master = master.sort_values("六維分", ascending=False)

    # 欄序
    front = [c for c in ["代號","名稱","評等","品質","分類","當前股價","最新價",
                          "外資90d淨","三大90d淨","外資持股Δpp","融資Δ%","借券90d量","籌碼分",
                          "1y報酬%","3y年化%","5y年化%","1y超額%","3y超額%",
                          "警戒筆數","六維分","六維訊號"] if c in master.columns]
    rest = [c for c in master.columns if c not in front]
    master = master[front + rest]

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(DST, engine="openpyxl") as xw:
        master.to_excel(xw, sheet_name="總覽", index=False)
        master.head(20).to_excel(xw, sheet_name="TOP20", index=False)
        master.tail(15).to_excel(xw, sheet_name="BOTTOM15", index=False)

    print(f"\n→ {DST}")
    print(f"\n=== TOP 15 ===")
    show = [c for c in ["代號","名稱","評等","品質","分類","籌碼分","1y報酬%","1y超額%","警戒筆數","六維分","六維訊號"] if c in master.columns]
    print(master[show].head(15).to_string(index=False))


if __name__ == "__main__":
    main()
