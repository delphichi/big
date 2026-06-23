# -*- coding: utf-8 -*-
"""
0050 納入預測引擎 — 在 ETF 強迫買入前先卡位
=======================================================================
邏輯:
  0050 純市值前 50 規則(FTSE),被納入 = 被動資金強迫買;但等 FTSE 公告/媒體報導,
  價格通常已先反映。本腳本在「市值排第 51-90 名 + 正在往上爬 + 基本面拐點 + 估值未爆」
  的時點就先標記出來,搶在 ETF 公告前進場。

資料 :
  data/台股財報估值.xlsx [財報估值比較] — 市值/收盤/PER/PBR/體質
  data/台股_體檢總表.xlsx [體檢總表]     — 評等/品質總分/估值/含金量/循環
  data/台股_拐點掃描.xlsx [全部訊號]     — 改善訊號數/分級(optional)

評分流程:
  1. 用「已知 0050 成分股」(在我們 universe 內)當 anchor → 算第 50 名門檻 ≈ anchors 最小市值
  2. 非 0050 候選 = 市值 ∈ [門檻 × 0.4, 門檻 × 1.5](即排名約 51-90 區)
  3. 加分:評等 A/B + 估值便宜或合理 + 拐點訊號 ≥ 2 + 含金量 ≥ 1.0
  4. 輸出 3 個分頁:候選/我們 universe 全市值排名/anchor 門檻參考

輸出:data/台股_0050納入預測.xlsx
"""
import os
import pandas as pd
import numpy as np

VAL = "data/台股財報估值.xlsx"
HEA = "data/台股_體檢總表.xlsx"
INF = "data/台股_拐點掃描.xlsx"
OUT = "data/台股_0050納入預測.xlsx"

# 已知 0050 成分股(2025 名單)— 此處只列我們 PICKS 內已知為 0050 成分的,當市值 anchor 用
# 不需完整 50 檔,只要有夠多筆能算最小值/中位數即可
KNOWN_0050 = {
    "2330",  # 台積電
    "2308",  # 台達電
    "2891",  # 中信金
    "2882",  # 國泰金
    "2881",  # 富邦金
    "2884",  # 玉山金
    "2885",  # 元大金
    "2880",  # 華南金
    "2883",  # 開發金
    "2887",  # 台新金
    "2412",  # 中華電
    "3045",  # 台灣大
    "2912",  # 統一超
    "3008",  # 大立光
    "6505",  # 台塑化
    "2379",  # 瑞昱
    "2395",  # 研華
    "2301",  # 光寶科
    "3711",  # 日月光投控
    "2890",  # 永豐金
    # 2026Q2 新納入
    "8046", "3443", "3665", "4958",
}


def main():
    if not os.path.exists(VAL):
        print(f"找不到 {VAL},請先跑 fetch_fundamentals_tw + backfill_market_cap"); return
    val = pd.read_excel(VAL, "財報估值比較")
    val["代號"] = val["代號"].astype(str)
    for c in ["市值(億)", "收盤", "PER(自算)", "PE位階%", "PBR", "殖利率%", "最新月營收年增%"]:
        if c in val.columns:
            val[c] = pd.to_numeric(val[c], errors="coerce")

    have_mcap = val.dropna(subset=["市值(億)"])
    print(f"有市值資料 {len(have_mcap)}/{len(val)} 檔")
    if have_mcap.empty:
        print("尚無市值資料 → 請先跑 backfill_market_cap.py"); return

    anchors = have_mcap[have_mcap["代號"].isin(KNOWN_0050)].sort_values("市值(億)")
    if anchors.empty:
        print("⚠️ 我們 universe 找不到 0050 anchor,無法估算門檻"); return
    anchor_min = float(anchors["市值(億)"].min())
    anchor_med = float(anchors["市值(億)"].median())
    print(f"0050 anchor {len(anchors)} 檔 → 最小市值 {anchor_min:.0f} 億 / 中位 {anchor_med:.0f} 億")
    print(f"  最小 anchor: {anchors.iloc[0]['代號']} {anchors.iloc[0]['名稱']} {anchor_min:.0f}億")

    threshold = anchor_min                          # 估算「第 50 名」門檻
    low  = threshold * 0.4                          # 約第 90 名左右
    high = threshold * 1.5                          # 比第 50 名稍大,捕捉剛卡邊的

    # 合併體檢
    hea = pd.DataFrame()
    if os.path.exists(HEA):
        hea = pd.read_excel(HEA, "體檢總表")
        hea["代號"] = hea["代號"].astype(str)
        hea = hea[["代號", "評等", "品質總分", "估值", "循環股", "含金量"]]
    # 合併拐點
    inf = pd.DataFrame()
    if os.path.exists(INF):
        try:
            inf = pd.read_excel(INF, "全部訊號")
            inf["代號"] = inf["代號"].astype(str)
            inf = inf[["代號", "改善訊號數", "分級"]]
        except Exception:
            pass

    base = have_mcap[["代號", "名稱", "市值(億)", "收盤", "PER(自算)", "PE位階%",
                      "PBR", "殖利率%", "最新月營收年增%"]].copy()
    if not hea.empty:
        base = base.merge(hea, on="代號", how="left")
    if not inf.empty:
        base = base.merge(inf, on="代號", how="left")

    base["是否0050"] = base["代號"].apply(lambda c: "✓" if c in KNOWN_0050 else "")
    base = base.sort_values("市值(億)", ascending=False).reset_index(drop=True)
    base["universe排名"] = base.index + 1

    # 候選:非 0050 + 市值落在 [low, high]
    cand = base[
        (base["是否0050"] == "") &
        (base["市值(億)"] >= low) &
        (base["市值(億)"] <= high)
    ].copy()

    # 評分:基本面強度 + 估值合理 + 拐點 + 動能
    def score(r):
        s = 0
        g = str(r.get("評等", ""))
        if g == "A": s += 30
        elif g == "B": s += 20
        elif g == "C": s += 10
        v = str(r.get("估值", ""))
        if "便宜" in v: s += 25
        elif "合理" in v: s += 18
        elif "偏貴" in v: s += 5
        ca = pd.to_numeric(r.get("含金量"), errors="coerce")
        if pd.notna(ca):
            if ca >= 1.2: s += 15
            elif ca >= 1.0: s += 10
            elif ca >= 0.8: s += 5
        sig = pd.to_numeric(r.get("改善訊號數"), errors="coerce")
        if pd.notna(sig):
            s += int(sig) * 5
        mo = pd.to_numeric(r.get("最新月營收年增%"), errors="coerce")
        if pd.notna(mo):
            if mo >= 30: s += 15
            elif mo >= 15: s += 10
            elif mo >= 0: s += 3
        # 越接近門檻越加分(已快進入區)
        gap = float(r["市值(億)"]) / threshold
        if 0.85 <= gap <= 1.0: s += 20   # 緊貼門檻
        elif 0.65 <= gap < 0.85: s += 12  # 看得到車尾燈
        elif 0.4 <= gap < 0.65: s += 5    # 還遠但有潛力
        return s

    cand["納入潛力分"] = cand.apply(score, axis=1)
    cand["距門檻%"] = (cand["市值(億)"] / threshold * 100).round(0)
    cand = cand.sort_values("納入潛力分", ascending=False)

    # 排序欄位
    out_cols = ["代號", "名稱", "市值(億)", "距門檻%", "納入潛力分", "評等", "品質總分",
                "估值", "改善訊號數", "分級", "含金量", "PER(自算)", "PE位階%",
                "PBR", "殖利率%", "最新月營收年增%", "循環股", "universe排名"]
    out_cols = [c for c in out_cols if c in cand.columns]

    anchors_view = anchors[["代號", "名稱", "市值(億)"]].copy()
    anchors_view = anchors_view.sort_values("市值(億)")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        cand[out_cols].to_excel(xw, sheet_name="納入候選", index=False)
        base[["universe排名", "代號", "名稱", "市值(億)", "是否0050",
              "評等", "估值", "PER(自算)", "PE位階%"]].to_excel(
            xw, sheet_name="universe市值排名", index=False)
        anchors_view.to_excel(xw, sheet_name="0050 anchor", index=False)
        # 門檻參考頁
        thresh_df = pd.DataFrame([
            {"項目": "Anchor 數", "值": len(anchors)},
            {"項目": "Anchor 最小市值(億)", "值": round(anchor_min, 1)},
            {"項目": "Anchor 中位市值(億)", "值": round(anchor_med, 1)},
            {"項目": "估算第50名門檻(億)", "值": round(threshold, 1)},
            {"項目": "候選下限(億)≈第90名", "值": round(low, 1)},
            {"項目": "候選上限(億)≈第40名", "值": round(high, 1)},
            {"項目": "候選檔數", "值": len(cand)},
        ])
        thresh_df.to_excel(xw, sheet_name="門檻說明", index=False)

    print(f"\n完成 → {OUT}")
    print(f"候選 {len(cand)} 檔(市值 {low:.0f}~{high:.0f} 億區)")
    if len(cand):
        top = cand.head(10)[["代號", "名稱", "市值(億)", "距門檻%", "納入潛力分",
                             "評等", "估值"]].to_string(index=False)
        print("\nTop 10 納入潛力:\n" + top)
    return cand


if __name__ == "__main__":
    main()
