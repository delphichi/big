# -*- coding: utf-8 -*-
"""
未來 PE 獨立報表 — 把「用預估明年 EPS 算的 PE」攤開
=======================================================================
讀 data/台股_體檢總表.xlsx,挑出有 forward 的(非循環非金融非資料不足),
按 PEG 排序,加上易讀的「兩把尺對照」與「四象限」分頁。

四象限(PE位階 vs PEG):
  🟢 過去便宜 + 未來便宜       = 雙低,最佳買點
  🔵 過去貴 + 未來便宜(PEG<1) = 成長卡位(本系統特色 alpha)
  ⚠️ 過去便宜 + 未來貴(PEG>2) = 成長耗盡/便宜陷阱
  🔴 過去貴 + 未來貴            = 真貴,不追

輸出:data/未來PE報表.xlsx(每日/每週自動更新隨體檢)
"""
import os
import pandas as pd

SRC = "data/台股_體檢總表.xlsx"
OUT = "data/未來PE報表.xlsx"


def quadrant(pe_pos, peg):
    pe_low = pd.notna(pe_pos) and pe_pos <= 30
    pe_high = pd.notna(pe_pos) and pe_pos >= 70
    peg_low = pd.notna(peg) and 0 < peg < 1
    peg_high = pd.notna(peg) and peg > 2
    if pe_low and peg_low:   return "🟢雙低(最佳)"
    if pe_high and peg_low:  return "🔵成長卡位"
    if pe_low and peg_high:  return "⚠️便宜陷阱"
    if pe_high and peg_high: return "🔴真貴"
    if peg_low:              return "🟢成長未反映"
    if peg_high:             return "⚠️成長耗盡"
    return "🟡中性"


def main():
    if not os.path.exists(SRC):
        print(f"找不到 {SRC},請先跑 stock_health_check.py"); return
    h = pd.read_excel(SRC, "體檢總表"); h["代號"] = h["代號"].astype(str)
    for c in ["PE位階", "PBR位階", "PEG", "ForwardPE", "ForwardPE保守", "成長率g%",
              "預估明年EPS", "PER", "ROE", "含金量", "月營收YoY", "EPS近3y%"]:
        if c in h.columns:
            h[c] = pd.to_numeric(h[c], errors="coerce")

    # 主表:有 forward 的(自動排除循環/金融/資料不足)
    fwd = h[h["未來估值"].notna()].copy()
    fwd["四象限"] = fwd.apply(lambda r: quadrant(r.get("PE位階"), r.get("PEG")), axis=1)
    fwd["PEG_s"] = fwd["PEG"]
    fwd = fwd.sort_values("PEG_s", na_position="last")

    cols = ["代號", "名稱", "評等", "品質總分", "四象限", "PER", "PE位階",
            "成長率g%", "預估明年EPS", "ForwardPE", "ForwardPE保守", "PEG",
            "估值", "未來估值", "ROE", "含金量", "月營收YoY", "EPS近3y%"]
    cols = [c for c in cols if c in fwd.columns]

    # 子表
    grade_ab = fwd[fwd["評等"].isin(["A", "B"])]
    growth_cap = fwd[fwd["四象限"] == "🔵成長卡位"]        # 過去貴但未來便宜(我們特色)
    double_low = fwd[fwd["四象限"] == "🟢雙低(最佳)"]
    trap = fwd[fwd["四象限"] == "⚠️便宜陷阱"]
    real_expensive = fwd[fwd["四象限"] == "🔴真貴"]

    # 循環股對照(沒 forward,但提示看 PBR)
    cyclical = h[h["循環股"].astype(str).str.contains("循環", na=False)][
        ["代號", "名稱", "評等", "PER", "PE位階", "PBR位階", "估值"]
    ].sort_values("PBR位階")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        fwd[cols].to_excel(xw, sheet_name="主表_按PEG排序", index=False)
        grade_ab[cols].to_excel(xw, sheet_name="A_B級_按PEG", index=False)
        double_low[cols].to_excel(xw, sheet_name="🟢雙低(最佳)", index=False)
        growth_cap[cols].to_excel(xw, sheet_name="🔵成長卡位", index=False)
        trap[cols].to_excel(xw, sheet_name="⚠️便宜陷阱", index=False)
        real_expensive[cols].to_excel(xw, sheet_name="🔴真貴", index=False)
        cyclical.to_excel(xw, sheet_name="循環股(看PBR非PE)", index=False)

        # 統計頁
        stat = pd.DataFrame([
            {"項目": "總體檢檔數", "值": len(h)},
            {"項目": "有 forward 的(非循環非金融)", "值": len(fwd)},
            {"項目": "🟢雙低(過去&未來都便宜)", "值": len(double_low)},
            {"項目": "🔵成長卡位(過去貴+PEG<1)", "值": len(growth_cap)},
            {"項目": "⚠️便宜陷阱(過去便宜+PEG>2)", "值": len(trap)},
            {"項目": "🔴真貴(過去&未來都貴)", "值": len(real_expensive)},
            {"項目": "循環股(看 PBR 不看 PE)", "值": len(cyclical)},
        ])
        stat.to_excel(xw, sheet_name="統計", index=False)

    print(f"完成 → {OUT}")
    print(f"  有 forward {len(fwd)} 檔 / A,B 級 {len(grade_ab)}")
    print(f"  🟢雙低 {len(double_low)} / 🔵成長卡位 {len(growth_cap)} / "
          f"⚠️陷阱 {len(trap)} / 🔴真貴 {len(real_expensive)}")
    if len(double_low):
        print("\n🟢雙低(最佳買點):")
        print(double_low.head(10)[["代號", "名稱", "評等", "PER", "PEG", "ForwardPE"]].to_string(index=False))
    if len(growth_cap):
        print("\n🔵成長卡位(過去貴但未來便宜):")
        print(growth_cap.head(10)[["代號", "名稱", "評等", "PER", "PEG", "ForwardPE"]].to_string(index=False))


if __name__ == "__main__":
    main()
