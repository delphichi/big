# -*- coding: utf-8 -*-
"""
台股 體檢總表 (Stock Health Check)
=======================================================================
依「找好公司的完整體檢框架 ①~⑩ + 循環股例外」對 data/台股財報估值.xlsx 全名單逐項打分。

品質總分(0~100) = 9 個面向加權(估值另計,因為『好公司≠好價格』):
  ⑥EPS真成長        20  (5年&近3年EPS CAGR;最重要的成長引擎,不是只看營收)
  ⑧獲利含金量        20  (OCF/淨利;賺的是不是真現金 — 最強照妖鏡)
  ⑨ROE              12  (資本運用效率)
  ②毛利率位階        10  (定價權/競爭力)
  ③營益率位階         8  (費用控管/營運槓桿)
  ④淨利率位階         8  (最終獲利能力)
  ⑤營收規模成長(5yCAGR) 8 (生意有沒有變大)
  ①營收動能(月YoY)    7  (最即時的成長訊號)
  ⑦EPS不落後營收      7  (抓股本稀釋/毛利漏水:營收長但EPS沒跟上)
估值(另計)        ⑩ PE位階 + PBR位階 → 便宜 / 合理 / 偏貴 / 過熱
循環股例外          逐年EPS曾為負 或 大起大落(max/min>3) → 標記,提示『PER失真,看PBR位階』
金融股(🏦)         毛利/營益口徑不適用 → 不評分,單獨標記

評等: A≥80  B 65-79  C 50-64  D<50
輸出: data/台股_體檢總表.xlsx
"""
import numpy as np, pandas as pd
from forward_pe import forward_metrics      # 單一真理來源(體檢/拐點/財報估值/0050預測共用)

SRC = "data/台股財報估值.xlsx"
OUT = "data/台股_體檢總表.xlsx"


def eps_cagr(eps_df, code):
    for k in eps_df.index:
        if str(k).split()[0] == code:
            v = [x for x in eps_df.loc[k].tolist() if pd.notna(x)]
            c5 = ((v[-1]/v[0])**(1/(len(v)-1))-1)*100 if len(v) >= 2 and v[0] > 0 and v[-1] > 0 else np.nan
            c3 = ((v[-1]/v[-3])**0.5-1)*100 if len(v) >= 3 and v[-3] > 0 and v[-1] > 0 else np.nan
            return c5, c3, v
    return np.nan, np.nan, []


def is_cyclical(v):
    """循環股 = 有真正的『衰退年』(EPS曾≤0,或某年YoY跌>20%),代表獲利上下震盪。
    純單調成長(年年往上,即使5年漲數倍)不算循環 — 那是結構性成長股。"""
    if len(v) < 3:
        return False
    if min(v) <= 0:
        return True
    return any(v[i] < v[i-1] * 0.8 for i in range(1, len(v)))   # 某年 EPS 較前年跌逾20%


def grade_quality(r):
    """回傳 (品質總分, 各分項dict, 主要漏洞list)。需先排除金融股。"""
    s, parts, leak = 0.0, {}, []

    # ⑥ EPS 真成長 (20)
    e5, e3 = r["EPS5y"], r["EPS3y"]
    if pd.notna(e5) and pd.notna(e3):
        if e5 >= 10 and e3 >= 10: p = 20
        elif e5 > 0 and e3 > 0:   p = 12
        elif (e5 < 0) ^ (e3 < 0): p = 4; leak.append("EPS單期衰退")
        else:                     p = 0; leak.append("EPS連年衰退")
    else:
        p = 0; leak.append("EPS資料不足")
    parts["⑥EPS成長"] = p; s += p

    # ⑧ 含金量 (20)
    g = r["獲利含金量"]
    if pd.isna(g):              p = 0; leak.append("無現金資料")
    elif g >= 1.2:             p = 20
    elif g >= 1.0:             p = 16
    elif g >= 0.7:             p = 10
    elif g >= 0.5:             p = 4;  leak.append(f"含金量{g:.1f}弱")
    else:                      p = 0;  leak.append(f"含金量{g:.1f}差")
    parts["⑧含金量"] = p; s += p

    # ⑨ ROE (12)
    roe = r["近四季ROE%"]
    if pd.isna(roe):           p = 0
    elif roe >= 20:            p = 12
    elif roe >= 15:            p = 9
    elif roe >= 12:            p = 6
    elif roe >= 8:             p = 3
    else:                      p = 0; leak.append(f"ROE{roe:.0f}低")
    parts["⑨ROE"] = p; s += p

    # ②③④ 三率位階
    def lvl(val, full, hi, mid):
        if pd.isna(val): return 0
        if val >= 60: return full
        if val >= 40: return hi
        if val >= 25: return mid
        return 0
    p = lvl(r["毛利率位階%"], 10, 7, 4); parts["②毛利位階"] = p; s += p
    if pd.notna(r["毛利率位階%"]) and r["毛利率位階%"] < 25: leak.append("毛利壓歷史低檔")
    p = lvl(r["營益率位階%"], 8, 6, 3); parts["③營益位階"] = p; s += p
    p = lvl(r["淨利率位階%"], 8, 6, 3); parts["④淨利位階"] = p; s += p

    # ⑤ 營收規模成長 (8)
    rc = r["5年營收CAGR%"]
    if pd.isna(rc):            p = 0
    elif rc >= 10:             p = 8
    elif rc >= 0:              p = 5
    else:                      p = 0; leak.append(f"營收5年萎縮{rc:.0f}%")
    parts["⑤營收成長"] = p; s += p

    # ① 營收動能 月YoY (7)
    mo = r["最新月營收年增%"]
    if pd.isna(mo):            p = 0
    elif mo >= 15:             p = 7
    elif mo >= 0:              p = 4
    else:                      p = 0; leak.append(f"月營收轉負{mo:.0f}%")
    parts["①營收動能"] = p; s += p

    # ⑦ EPS 不落後營收 (7) — 抓稀釋/毛利漏水
    if pd.notna(e5) and pd.notna(rc) and rc > 0:
        if e5 >= rc:           p = 7
        elif e5 >= 0.5*rc:     p = 4
        else:                  p = 0; leak.append("EPS遠落後營收(稀釋/毛利漏)")
    elif pd.notna(e5) and pd.notna(rc) and rc <= 0:
        p = 4 if e5 > 0 else 0   # 營收沒長,EPS有長(靠效率)也給點
    else:
        p = 0
    parts["⑦EPS跟上營收"] = p; s += p

    # ⑪ 動態惡化扣分(最高扣 -15)— 抓「當下漂亮但正在變差」的陷阱
    #   核心競爭力流失(ROE 從 5年均>=15 滑到 <67%)→ 扣 10
    #   短期償債警報(負債>70 + 流動<100)→ 扣 10
    #   高槓桿(負債>80,即使流動還OK)→ 扣 5
    penalty = 0
    roe_cur = r.get("近四季ROE%"); roe_avg = r.get("ROE5年均")
    if pd.notna(roe_cur) and pd.notna(roe_avg) and roe_avg >= 15 and roe_cur < roe_avg * 0.67:
        penalty += 10
        leak.append(f"⚠️ROE滑落({roe_cur:.0f}<5年均{roe_avg:.0f}×67%)")
    dr = r.get("負債比%"); cr = r.get("流動比%")
    if pd.notna(dr) and dr > 70 and pd.notna(cr) and cr < 100:
        penalty += 10
        leak.append(f"⚠️短期償債警報(負債{dr:.0f}+流動{cr:.0f})")
    elif pd.notna(dr) and dr > 80:
        penalty += 5
        leak.append(f"⚠️高槓桿(負債{dr:.0f}%)")
    # 存貨爆衝 vs 營收背離(八方案例:存貨+38.8 vs 營收+3.9)→ 需求軟警訊
    inv_yoy = r.get("存貨年增%"); rev_yoy = r.get("最新月營收年增%")
    if pd.notna(inv_yoy) and inv_yoy >= 40 and (pd.isna(rev_yoy) or rev_yoy < inv_yoy * 0.3):
        penalty += 5
        leak.append(f"⚠️存貨爆衝(存{inv_yoy:.0f}%vs營{rev_yoy if pd.notna(rev_yoy) else 0:.0f}%)")
    penalty = min(penalty, 15)              # 累計上限 15(避免雙條都中扣到評等崩盤)
    parts["⑪動態惡化扣分"] = -penalty
    s -= penalty

    return round(s, 1), parts, leak


def valuation_tag(pe, pbr):
    """⑩ 估值標籤:綜合 PE位階 + PBR位階。"""
    xs = [x for x in (pe, pbr) if pd.notna(x)]
    if not xs: return "—"
    m = np.mean(xs)
    if m <= 30: return "🟢便宜"
    if m <= 55: return "🟡合理"
    if m <= 80: return "🟠偏貴"
    return "🔴過熱"


def main():
    val = pd.read_excel(SRC, "財報估值比較"); val["代號"] = val["代號"].astype(str)
    his = pd.read_excel(SRC, "相對歷史水位"); his["代號"] = his["代號"].astype(str)
    eps = pd.read_excel(SRC, "逐年EPS", index_col=0)

    m = val.merge(his[["代號", "毛利率位階%", "營益率位階%", "淨利率位階%", "ROE位階%",
                       "PER位階%", "PBR位階%", "ROE5年均"]],
                  on="代號", how="left")
    for c in ["近四季ROE%", "獲利含金量", "5年營收CAGR%", "最新月營收年增%", "近四季EPS",
              "PER(自算)", "PE位階%", "毛利率位階%", "營益率位階%", "淨利率位階%", "PBR位階%",
              "殖利率%", "收盤", "負債比%", "流動比%", "存貨年增%", "營益率%", "ROE5年均",
              "合理價", "偏便宜價", "深度買點價"]:
        if c in m.columns:
            m[c] = pd.to_numeric(m[c], errors="coerce")

    rows = []
    for _, r in m.iterrows():
        c = r["代號"]
        e5, e3, v = eps_cagr(eps, c)
        rr = dict(r); rr["EPS5y"], rr["EPS3y"] = e5, e3
        cyc = is_cyclical(v)
        fin = pd.notna(r.get("金融"))
        if fin:
            score, parts, leak, grade = np.nan, {}, ["金融股:口徑不適用,不評分"], "金融🏦"
        else:
            score, parts, leak = grade_quality(rr)
            grade = "A" if score >= 80 else "B" if score >= 65 else "C" if score >= 50 else "D"
        # ROIC 近似:用 ROE × (1 - 負債比) → 把槓桿剝掉,接近「無槓桿資本回報」
        # 比真實 ROIC 保守(沒考慮稅後營業利益和投入資本細節),但能分辨「靠槓桿」vs「真實效率」
        # 注意:真實 ROIC 應該是 NOPAT / 投入資本(權益+長債),需 raw 財報重算,下版納入
        roic_proxy = None
        if pd.notna(r["近四季ROE%"]) and pd.notna(r.get("負債比%")):
            roic_proxy = round(float(r["近四季ROE%"]) * (1 - float(r["負債比%"]) / 100), 1)
        out = {"代號": c, "名稱": r["名稱"], "評等": grade, "品質總分": score,
               "EPS5y%": round(e5, 1) if pd.notna(e5) else None,
               "EPS近3y%": round(e3, 1) if pd.notna(e3) else None,
               "ROE": r["近四季ROE%"], "ROE5年均": r.get("ROE5年均"),
               "ROIC估算": roic_proxy, "含金量": r["獲利含金量"],
               "負債比%": r.get("負債比%"), "流動比%": r.get("流動比%"),
               "存貨年增%": r.get("存貨年增%"),
               "毛利位階": r["毛利率位階%"], "淨利位階": r["淨利率位階%"],
               "營收5yCAGR": r["5年營收CAGR%"], "月營收YoY": r["最新月營收年增%"],
               "PER": round(r["PER(自算)"], 1) if pd.notna(r["PER(自算)"]) else None,
               "PE位階": r["PE位階%"], "PBR位階": r["PBR位階%"],
               "估值": valuation_tag(r["PE位階%"], r["PBR位階%"]),
               "殖利率": r["殖利率%"],
               "循環股": "⚠️循環(看PBR)" if cyc else "",
               "主要漏洞": "、".join(leak[:3])}
        # Forward PE:用未來 EPS 看現價(股票買的是未來不是過去)
        fwd = forward_metrics(r.get("收盤"), r.get("近四季EPS"), r.get("PER(自算)"),
                              e3, e5, r.get("最新月營收年增%"), cyc)
        out.update(fwd)
        # 合理價鬧鐘:直接帶過(財報估值已用「歷史PE分位 × forward EPS」算好)
        for k in ("合理價", "偏便宜價", "深度買點價"):
            out[k] = r.get(k)
        # 現價 vs 鬧鐘:離哪一層最近?幫你免心算
        close = r.get("收盤")
        if pd.notna(close) and pd.notna(out.get("合理價")):
            if close <= out["深度買點價"]:    out["鬧鐘"] = "💎深度買點"
            elif close <= out["偏便宜價"]:    out["鬧鐘"] = "🟢偏便宜"
            elif close <= out["合理價"]:      out["鬧鐘"] = "🟡合理"
            else:                              out["鬧鐘"] = "🔴貴於合理價"
        out.update(parts)
        rows.append(out)

    df = pd.DataFrame(rows)
    nonfin = df[df["評等"] != "金融🏦"].copy()
    df = pd.concat([nonfin.sort_values("品質總分", ascending=False),
                    df[df["評等"] == "金融🏦"]], ignore_index=True)

    part_cols = ["⑥EPS成長", "⑧含金量", "⑨ROE", "②毛利位階", "③營益位階",
                 "④淨利位階", "⑤營收成長", "①營收動能", "⑦EPS跟上營收",
                 "⑪動態惡化扣分"]
    base = ["代號", "名稱", "評等", "品質總分", "EPS5y%", "EPS近3y%",
            "ROE", "ROE5年均", "ROIC估算", "負債比%", "流動比%", "存貨年增%", "含金量",
            "毛利位階", "淨利位階", "營收5yCAGR", "月營收YoY", "PER", "PE位階", "PBR位階",
            "估值", "成長率g%", "預估明年EPS", "ForwardPE", "ForwardPE保守", "PEG", "未來估值",
            "鬧鐘", "合理價", "偏便宜價", "深度買點價",
            "殖利率", "循環股", "主要漏洞"]
    for col in ["成長率g%", "預估明年EPS", "ForwardPE", "ForwardPE保守", "PEG", "未來估值"]:
        if col not in df.columns:
            df[col] = None
    full = df[base + part_cols]

    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        full.to_excel(xw, sheet_name="體檢總表", index=False)
        a = nonfin[nonfin["評等"] == "A"].sort_values("品質總分", ascending=False)
        a[base].to_excel(xw, sheet_name="A級好公司", index=False)
        ace = a[a["估值"].isin(["🟢便宜", "🟡合理"])]
        ace[base].to_excel(xw, sheet_name="A級+好價格", index=False)
        cyc = nonfin[nonfin["循環股"] != ""].sort_values("PBR位階")
        cyc[base].to_excel(xw, sheet_name="循環股(看PBR)", index=False)
        # 未來估值便宜的成長股:Forward PE 看起來便宜 + 評等不差(用未來修正當下 PE)
        fwd = nonfin[nonfin["未來估值"].isin(["🟢成長未反映", "🟢未來便宜", "🟡未來合理"])
                     & nonfin["評等"].isin(["A", "B"])].copy()
        fwd = fwd.sort_values(["PEG", "ForwardPE"], na_position="last")
        fwd[base].to_excel(xw, sheet_name="未來估值便宜(ForwardPE)", index=False)

    print(f"完成 → {OUT}  (評分 {len(nonfin)} 檔 + 金融 {len(df)-len(nonfin)} 檔)")
    print("評等分布:", nonfin["評等"].value_counts().reindex(["A","B","C","D"]).to_dict())
    print(f"A級好公司 {len(nonfin[nonfin['評等']=='A'])} 檔 / 其中估值便宜或合理 "
          f"{len(nonfin[(nonfin['評等']=='A') & (nonfin['估值'].isin(['🟢便宜','🟡合理']))])} 檔")
    return full


if __name__ == "__main__":
    main()
