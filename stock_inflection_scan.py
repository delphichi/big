# -*- coding: utf-8 -*-
"""
拐點掃描 (Inflection Scanner) — 第二層『變化率』:抓「正在變好、市場還沒反映」的便宜公司
=======================================================================
體檢總表抓「已經好」(水準);這支抓「正在變好」(變化率的二階導數)。
Howard Marks 二階思考:超額報酬來自在『市場還沒認可它是好公司、估值還低』時就抓到拐點。

讀 data/台股財報估值.xlsx,對每檔(排除金融)算「改善訊號」:
  A. 毛利率拐頭↑   : 最新季毛利 > 前4季均(需逐季資料;無則標『待逐季』)
  B. 月營收動能    : 最新月營收年增% ≥ 10(>0 至少正)
  C. EPS 加速      : EPS 近3年CAGR > EPS 5年CAGR(近年比長期快=加速)
  D. ROE 回升      : 近四季ROE > 5年平均ROE(獲利效率改善)
  現金門檻         : 獲利含金量 ≥ 0.8(別抓到帳面改善、現金沒跟上的)
  估值仍低(關鍵)   : 非循環看 PE位階≤50;循環看 PBR位階≤50(市場還沒再評價,才有 edge)

拐點候選 = 估值仍低 + 含金量OK + 改善訊號≥2 + EPS未連年衰退。
分級:🔥強拐點(訊號≥3) / 🌱初拐點(=2) / —(不足)。輸出 data/台股_拐點掃描.xlsx。
"""
import os
import numpy as np
import pandas as pd

SRC = "data/台股財報估值.xlsx"
OUT = "data/台股_拐點掃描.xlsx"


def eps_series(eps_df, code):
    for k in eps_df.index:
        if str(k).split()[0] == code:
            return [x for x in eps_df.loc[k].tolist() if pd.notna(x)]
    return []


def cagr(v, n):
    if len(v) >= n + 1 and v[-(n+1)] > 0 and v[-1] > 0:
        return (v[-1] / v[-(n+1)]) ** (1 / n) - 1
    return np.nan


def load_quarterly_margin():
    """若財報估值.xlsx 有『逐季毛利率』分頁(未來逐季資料補上後)就讀,回 {代號:[季毛利...]};否則回 {}。"""
    try:
        q = pd.read_excel(SRC, "逐季毛利率", index_col=0)
        out = {}
        for k in q.index:
            code = str(k).split()[0]
            out[code] = [x for x in q.loc[k].tolist() if pd.notna(x)]
        return out
    except Exception:
        return {}


def main():
    val = pd.read_excel(SRC, "財報估值比較"); val["代號"] = val["代號"].astype(str)
    his = pd.read_excel(SRC, "相對歷史水位"); his["代號"] = his["代號"].astype(str)
    eps = pd.read_excel(SRC, "逐年EPS", index_col=0)
    qm = load_quarterly_margin()
    has_q = len(qm) > 0

    m = val.merge(his[["代號", "PER位階%", "PBR位階%", "毛利率位階%"]], on="代號", how="left")
    for c in ["近四季ROE%", "5年平均ROE%", "獲利含金量", "PE位階%", "PBR位階%",
              "最新月營收年增%", "近四季EPS", "5年營收CAGR%"]:
        m[c] = pd.to_numeric(m[c], errors="coerce")
    nf = m[m["金融"].isna()].copy()

    rows = []
    for _, r in nf.iterrows():
        c = r["代號"]
        v = eps_series(eps, c)
        e5, e3 = cagr(v, min(4, len(v)-1)) if len(v) >= 2 else np.nan, cagr(v, 3) if len(v) >= 4 else np.nan
        e5p, e3p = (e5*100 if pd.notna(e5) else np.nan), (e3*100 if pd.notna(e3) else np.nan)
        cyclical = (len(v) >= 3) and (min(v) <= 0 or any(v[i] < v[i-1]*0.8 for i in range(1, len(v))))

        # --- 改善訊號 ---
        # A 毛利拐頭(需逐季)
        if has_q and c in qm and len(qm[c]) >= 5:
            qs = qm[c]; A = qs[-1] > np.mean(qs[-5:-1])
            A_txt = "✓" if A else "✗"
        else:
            A = False; A_txt = "待逐季"
        # B 月營收動能
        mo = r["最新月營收年增%"]
        B = pd.notna(mo) and mo >= 10
        # C EPS 加速
        C = pd.notna(e3p) and pd.notna(e5p) and e3p > e5p and e3p > 0
        # D ROE 回升
        roe, roe5 = r["近四季ROE%"], r["5年平均ROE%"]
        D = pd.notna(roe) and pd.notna(roe5) and roe > roe5
        sig = sum([A, B, C, D])

        g = r["獲利含金量"]
        cash_ok = pd.notna(g) and g >= 0.8
        # 估值仍低:循環看PBR,其餘看PE
        vpos = r["PBR位階%"] if cyclical else r["PE位階%"]
        cheap = pd.notna(vpos) and vpos <= 50
        eps_not_dying = not (pd.notna(e5p) and pd.notna(e3p) and e5p < 0 and e3p < 0)

        candidate = cheap and cash_ok and eps_not_dying and sig >= 2
        tier = ("🔥強拐點" if sig >= 3 else "🌱初拐點") if candidate else ""

        rows.append({
            "代號": c, "名稱": r["名稱"], "分級": tier, "改善訊號數": sig,
            "A毛利拐頭": A_txt, "B月營收動能": "✓" if B else "✗",
            "C_EPS加速": "✓" if C else "✗", "D_ROE回升": "✓" if D else "✗",
            "EPS5y%": round(e5p, 1) if pd.notna(e5p) else None,
            "EPS近3y%": round(e3p, 1) if pd.notna(e3p) else None,
            "近四季ROE": roe, "5年均ROE": roe5, "月營收YoY": mo,
            "含金量": g, "循環": "⚠️" if cyclical else "",
            "看哪個位階": "PBR" if cyclical else "PE",
            "估值位階": vpos, "PER": round(r["PER(自算)"], 1) if pd.notna(r["PER(自算)"]) else None,
            "殖利率%": r.get("殖利率%"),
        })

    df = pd.DataFrame(rows)
    cand = df[df["分級"] != ""].sort_values(["改善訊號數", "估值位階"], ascending=[False, True])
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        cand.to_excel(xw, sheet_name="潛在拐點觀察區", index=False)
        df.sort_values("改善訊號數", ascending=False).to_excel(xw, sheet_name="全部訊號", index=False)

    print(f"完成 → {OUT}  (逐季毛利{'已接' if has_q else '待補,A訊號暫無'})")
    print(f"潛在拐點候選 {len(cand)} 檔:🔥強 {len(cand[cand['分級']=='🔥強拐點'])} / 🌱初 {len(cand[cand['分級']=='🌱初拐點'])}")
    return cand


if __name__ == "__main__":
    main()
