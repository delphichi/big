# -*- coding: utf-8 -*-
"""
0050 納入預測引擎 — 在 ETF 強迫買入前先卡位
=======================================================================
邏輯:
  0050 純市值前 50 規則(FTSE),被納入 = 被動資金強迫買;但等 FTSE 公告/媒體報導,
  價格通常已先反映。本腳本在「市值排第 51-90 名 + 正在往上爬 + 基本面拐點 + 估值未爆」
  的時點就先標記出來,搶在 ETF 公告前進場。

資料 :
  data/twse_top100_marketcap.csv  — 全市場前 100 大實際市值排名(人工 anchor,
                                    供精確估算第 50 名門檻;若無則退回 anchor 估算)
  data/台股財報估值.xlsx [財報估值比較] — 我們 universe 的市值/PER/PBR/體質
  data/台股_體檢總表.xlsx [體檢總表]     — 評等/品質總分/估值/含金量/循環
  data/台股_拐點掃描.xlsx [全部訊號]     — 改善訊號數/分級(optional)

評分流程:
  1. 從 twse_top100 取第 50 名市值當實際 0050 門檻
  2. 候選 = 排名 51-90(市值在 [第90名, 第50名*1.05] 區)且不在 0050 內
  3. 加分:評等 A/B + 估值便宜或合理 + 拐點訊號 ≥ 2 + 含金量 ≥ 1.0
  4. 兩種候選來源並列:
     (a) 我們 universe 內 → 已有完整體檢,可直接評分
     (b) 不在我們 universe → blind spot,標記為「該加進 PICKS」
  5. 輸出 5 分頁:候選A(已體檢) / 候選B(blind spot) / universe市值排名 /
                  0050 anchor / 門檻說明

輸出:data/台股_0050納入預測.xlsx
"""
import os
import pandas as pd
import numpy as np

VAL = "data/台股財報估值.xlsx"
HEA = "data/台股_體檢總表.xlsx"
INF = "data/台股_拐點掃描.xlsx"
TOP = "data/twse_top100_marketcap.csv"
OUT = "data/台股_0050納入預測.xlsx"

# 已知 0050 成分股(2025 名單)— 用於 universe 識別,不再用作門檻估算 anchor
KNOWN_0050 = {
    "2330", "2454", "2317", "2308", "3711", "2382", "2412", "3045", "2912",
    "3008", "6505", "2207", "2379", "2395", "2301", "2002", "1303", "1301",
    "2891", "2882", "2881", "2884", "2885", "2880", "2883", "2887", "2890",
    "2886", "2892", "5880", "2801", "2885", "5871",
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

    # ---- 真實門檻:讀 twse_top100 ----
    if os.path.exists(TOP):
        top = pd.read_csv(TOP, dtype={"代號": str})
        threshold = float(top[top["rank"] == 50]["市值億"].iloc[0])   # 第 50 名
        rank90   = float(top[top["rank"] == 90]["市值億"].iloc[0])   # 第 90 名
        rank60   = float(top[top["rank"] == 60]["市值億"].iloc[0])   # 第 60 名
        rank100  = float(top[top["rank"] == 100]["市值億"].iloc[0])
        top100_set = set(top["代號"])
        print(f"實際 0050 門檻(rank 50): {threshold:.0f}億 / rank 60: {rank60:.0f} / rank 90: {rank90:.0f} / rank 100: {rank100:.0f}")
        src = "twse_top100"
        candidate_low = rank90 * 0.9     # 略放寬到第 95 名
        candidate_high = threshold * 1.05   # 緊貼門檻上方也納入(剛卡邊)
    else:
        # Fallback: anchor 估算
        have = val.dropna(subset=["市值(億)"])
        anchors = have[have["代號"].isin(KNOWN_0050)].sort_values("市值(億)")
        threshold = float(anchors["市值(億)"].min()) if len(anchors) else 4000
        candidate_low, candidate_high = threshold * 0.4, threshold * 1.5
        top, top100_set = pd.DataFrame(), set()
        print(f"⚠️ {TOP} 不存在,改用 anchor 估算門檻 {threshold:.0f}億")
        src = "anchor"

    have_mcap = val.dropna(subset=["市值(億)"])
    print(f"我們 universe 有市值資料 {len(have_mcap)}/{len(val)} 檔")

    # ---- 合併體檢/拐點 ----
    hea = pd.DataFrame()
    if os.path.exists(HEA):
        hea = pd.read_excel(HEA, "體檢總表")
        hea["代號"] = hea["代號"].astype(str)
        hea = hea[["代號", "評等", "品質總分", "估值", "循環股", "含金量"]]
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

    # ---- 候選 A:我們 universe 內 + 在 51-90 區 ----
    cand_a = base[
        (base["是否0050"] == "") &
        (base["市值(億)"] >= candidate_low) &
        (base["市值(億)"] <= candidate_high)
    ].copy()

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
        gap = float(r["市值(億)"]) / threshold
        if 0.85 <= gap <= 1.05: s += 20
        elif 0.65 <= gap < 0.85: s += 12
        elif gap < 0.65: s += 5
        return s

    if len(cand_a):
        cand_a["納入潛力分"] = cand_a.apply(score, axis=1)
        cand_a["距門檻%"] = (cand_a["市值(億)"] / threshold * 100).round(0)
        cand_a = cand_a.sort_values("納入潛力分", ascending=False)

    # ---- 候選 B:blind spot — 在 twse_top100 rank 51-90 但不在我們 PICKS ----
    cand_b = pd.DataFrame()
    if src == "twse_top100":
        our_codes = set(val["代號"].astype(str))
        spot = top[(top["rank"] >= 51) & (top["rank"] <= 90)
                   & (~top["代號"].isin(KNOWN_0050))
                   & (~top["代號"].isin(our_codes))].copy()
        spot["距門檻%"] = (spot["市值億"] / threshold * 100).round(0)
        spot["建議"] = "加入 PICKS 抓體檢"
        cand_b = spot[["rank", "代號", "名稱", "市值億", "距門檻%", "建議"]]

    # ---- 輸出 ----
    out_cols = ["代號", "名稱", "市值(億)", "距門檻%", "納入潛力分", "評等", "品質總分",
                "估值", "改善訊號數", "分級", "含金量", "PER(自算)", "PE位階%",
                "PBR", "殖利率%", "最新月營收年增%", "循環股", "universe排名"]
    out_cols = [c for c in out_cols if c in cand_a.columns]

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        cand_a[out_cols].to_excel(xw, sheet_name="候選A_已體檢", index=False)
        cand_b.to_excel(xw, sheet_name="候選B_blind_spot", index=False)
        base[["universe排名", "代號", "名稱", "市值(億)", "是否0050", "評等",
              "估值", "PER(自算)", "PE位階%"]].to_excel(
            xw, sheet_name="universe市值排名", index=False)
        if not top.empty:
            top.to_excel(xw, sheet_name="實際前100大", index=False)
        thresh_df = pd.DataFrame([
            {"項目": "資料來源", "值": src},
            {"項目": "0050 門檻(rank 50, 億)", "值": round(threshold, 1)},
            {"項目": "候選下限(億)", "值": round(candidate_low, 1)},
            {"項目": "候選上限(億)", "值": round(candidate_high, 1)},
            {"項目": "候選A (已體檢) 檔數", "值": len(cand_a)},
            {"項目": "候選B (blind spot) 檔數", "值": len(cand_b)},
        ])
        thresh_df.to_excel(xw, sheet_name="門檻說明", index=False)

    print(f"\n完成 → {OUT}")
    print(f"候選A (已體檢) {len(cand_a)} 檔 / 候選B (blind spot) {len(cand_b)} 檔")
    if len(cand_a):
        top_show = cand_a.head(10)[["代號", "名稱", "市值(億)", "距門檻%", "納入潛力分",
                                     "評等", "估值"]].to_string(index=False)
        print(f"\n候選A Top 10:\n{top_show}")
    if len(cand_b):
        print(f"\n候選B (top100 但不在 PICKS):\n{cand_b.head(20).to_string(index=False)}")
    return cand_a


if __name__ == "__main__":
    main()
