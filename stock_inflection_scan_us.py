# -*- coding: utf-8 -*-
"""
美股拐點掃描 (US Inflection Scanner) — 沿用台股『正在變好的便宜公司』邏輯,改用美股欄位
=======================================================================
讀 data/美股體檢總表.xlsx (由 fetch_fundamentals_us.py 產出),對非循環非衰退股算改善訊號:
  A. 毛利率高且穩(>40%,美股年報沒逐季,用絕對水準代理)
  B. 營收CAGR>5  (有成長動能)
  C. EPS 加速   (EPS3y > EPS5y 且 >0,近期比長期快)
  D. ROE 改善代理(ROE>=15;美股無 5年均 ROE 可比,用絕對門檻)
  含金量門檻   : ≥ 0.8 (賺的是真現金)
  估值仍低(關鍵): 非循環看 PER ≤ 22(美股無歷史位階,用絕對倍數)
                  循環看 P/B ≤ 2.5

候選 = 估值仍低 + 含金量OK + 改善訊號≥2 + EPS未連年衰退。輸出 data/美股_拐點掃描.xlsx。
"""
import os
import numpy as np
import pandas as pd

SRC = "data/美股體檢總表.xlsx"
OUT = "data/美股_拐點掃描.xlsx"


def main():
    if not os.path.exists(SRC):
        print(f"找不到 {SRC},請先跑 fetch_fundamentals_us.py(us-fundamentals workflow)"); return
    df = pd.read_excel(SRC, "體檢總表")
    for c in ["品質總分", "EPS5y%", "EPS3y%", "ROE%", "含金量", "毛利率%", "淨利率%",
              "營收CAGR%", "PER", "PBR", "PEG", "殖利率%"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    rows = []
    for _, r in df.iterrows():
        cyclical = "循環" in str(r.get("循環股", ""))
        # 改善訊號
        A = pd.notna(r["毛利率%"]) and r["毛利率%"] >= 40       # 美股年報,用毛利水準代理拐頭
        B = pd.notna(r["營收CAGR%"]) and r["營收CAGR%"] >= 5
        C = pd.notna(r["EPS3y%"]) and pd.notna(r["EPS5y%"]) and r["EPS3y%"] > r["EPS5y%"] and r["EPS3y%"] > 0
        D = pd.notna(r["ROE%"]) and r["ROE%"] >= 15
        sig = sum([A, B, C, D])

        cash_ok = pd.notna(r["含金量"]) and r["含金量"] >= 0.8
        if cyclical:
            cheap = pd.notna(r["PBR"]) and r["PBR"] > 0 and r["PBR"] <= 2.5
            v_metric = f"PBR {r['PBR']}"
        else:
            cheap = pd.notna(r["PER"]) and r["PER"] > 0 and r["PER"] <= 22
            v_metric = f"PER {r['PER']}"
        eps_not_dying = not (pd.notna(r["EPS5y%"]) and pd.notna(r["EPS3y%"]) and r["EPS5y%"] < 0 and r["EPS3y%"] < 0)

        candidate = cheap and cash_ok and eps_not_dying and sig >= 2
        tier = ("🔥強拐點" if sig >= 3 else "🌱初拐點") if candidate else ""

        rows.append({
            "代號": r["代號"], "名稱": r.get("名稱"), "產業": r.get("產業"),
            "分級": tier, "改善訊號數": sig,
            "A毛利≥40": "✓" if A else "✗", "B營收CAGR≥5": "✓" if B else "✗",
            "C_EPS加速": "✓" if C else "✗", "D_ROE≥15": "✓" if D else "✗",
            "EPS5y%": r["EPS5y%"], "EPS近3y%": r["EPS3y%"],
            "ROE%": r["ROE%"], "毛利率%": r["毛利率%"], "含金量": r["含金量"],
            "PER": r["PER"], "PBR": r["PBR"], "PEG": r.get("PEG"),
            "看哪個估值": "PBR(循環)" if cyclical else "PER",
            "估值評語": v_metric, "循環": "⚠️" if cyclical else "",
            "殖利率%": r.get("殖利率%"), "市值(億美)": r.get("市值(億美)"),
        })

    out = pd.DataFrame(rows)
    cand = out[out["分級"] != ""].sort_values(["改善訊號數", "PER"], ascending=[False, True])
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        cand.to_excel(xw, sheet_name="潛在拐點觀察區", index=False)
        out.sort_values("改善訊號數", ascending=False).to_excel(xw, sheet_name="全部訊號", index=False)

    print(f"完成 → {OUT}")
    print(f"美股拐點 {len(cand)} 檔:🔥強 {len(cand[cand['分級']=='🔥強拐點'])} / 🌱初 {len(cand[cand['分級']=='🌱初拐點'])}")
    return cand


if __name__ == "__main__":
    main()
