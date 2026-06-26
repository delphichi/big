# -*- coding: utf-8 -*-
"""
美股體檢補充欄(本地計算, 不打 FMP) patch_us_addons.py
=======================================================================
從現有「美股體檢總表.xlsx」欄位**衍生**出三個原本缺的判讀:
  ⑦ EPS不落後營收  — 抓股本稀釋/毛利漏水
  ⑪ ROE滑落扣分     — 當前 ROE 顯著低於 5 年均 = 走下坡
  主要漏洞清單      — 把已知問題彙整成一欄, 逐檔分析時一眼看出

不再次抓 FMP, 純讀現有欄位算 → 秒級完成。
"""
import os
import pandas as pd
import numpy as np

SRC = "data/美股體檢總表.xlsx"


def calc_eps_lag_revenue(r):
    """⑦ EPS 不落後營收:營收成長強(≥10%)但 EPS 成長遠落後(<營收的 50%)= 稀釋/漏水"""
    rev = r.get("營收CAGR%"); eps3 = r.get("EPS3y%")
    if pd.isna(rev) or pd.isna(eps3) or rev < 10:
        return ""    # 營收弱不適用
    if eps3 < rev * 0.5:
        return f"⚠️EPS落後營收(EPS3y {eps3:.0f}<營收CAGR {rev:.0f}×50%)"
    return ""


def calc_roe_decay(r):
    """⑪ ROE 滑落:當前 ROE < 5 年均 × 67% 且 5 年均 ≥ 15(原本是好公司,正在轉弱)"""
    cur = r.get("ROE%"); avg = r.get("ROE5年均%")
    if pd.isna(cur) or pd.isna(avg) or avg < 15:
        return 0, ""
    if cur < avg * 0.67:
        return -10, f"⚠️ROE滑落({cur:.0f}<5年均{avg:.0f}×67%)"
    if cur < avg * 0.8:
        return -5, f"🟡ROE弱化({cur:.0f}<5年均{avg:.0f}×80%)"
    return 0, ""


def calc_debt_alarm(r):
    """⑪ 短期償債警報(等負債比 patch 跑完才有效;現在先預埋邏輯)
       負債比 > 70 + 流動比 < 1.0 = 短期償債吃緊"""
    debt = r.get("負債比%"); cur = r.get("流動比")
    if pd.isna(debt) or pd.isna(cur):
        return 0, ""
    # 美股流動比是小數(1.5=150%), 轉百分比口徑
    cur_pct = cur * 100 if cur < 10 else cur
    if debt > 80:
        return -10, f"⚠️高槓桿(負債{debt:.0f}%)"
    if debt > 70 and cur_pct < 100:
        return -10, f"⚠️短期償債警報(負債{debt:.0f}+流動{cur_pct:.0f})"
    return 0, ""


def calc_cash_quality_alarm(r):
    """⑧ 含金量 < 0.5 = 帳面獲利沒收到現金(警訊)"""
    c = r.get("含金量")
    if pd.isna(c) or c >= 0.5:
        return 0, ""
    if c < 0:
        return -10, f"⚠️現金流為負(含金量{c:.2f})"
    return -5, f"🟡現金品質差(含金量{c:.2f}<0.5)"


def main():
    xls = pd.ExcelFile(SRC)
    sheets = {sh: pd.read_excel(SRC, sheet_name=sh) for sh in xls.sheet_names}
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    for c in ["營收CAGR%", "EPS3y%", "EPS5y%", "ROE%", "ROE5年均%",
              "含金量", "負債比%", "流動比", "毛利率%", "淨利率%"]:
        if c in h.columns:
            h[c] = pd.to_numeric(h[c], errors="coerce")

    print(f"處理 {len(h)} 檔...")

    # 計算各補充欄
    h["⑦EPS落後營收"] = h.apply(calc_eps_lag_revenue, axis=1)

    roe_pen, roe_msg = [], []
    debt_pen, debt_msg = [], []
    cash_pen, cash_msg = [], []
    for _, r in h.iterrows():
        p, m = calc_roe_decay(r);          roe_pen.append(p); roe_msg.append(m)
        p, m = calc_debt_alarm(r);         debt_pen.append(p); debt_msg.append(m)
        p, m = calc_cash_quality_alarm(r); cash_pen.append(p); cash_msg.append(m)
    h["⑪ROE滑落"] = roe_msg
    h["⑪償債警報"] = debt_msg
    h["⑧現金品質警報"] = cash_msg

    # ⑪ 動態惡化扣分總和(上限 -15)
    pen_total = pd.Series(roe_pen) + pd.Series(debt_pen) + pd.Series(cash_pen)
    pen_total = pen_total.clip(lower=-15)
    h["⑪動態惡化扣分"] = pen_total.values

    # 主要漏洞清單(把所有警報串起來)
    def build_leak(r):
        leaks = []
        for col in ("⑦EPS落後營收", "⑪ROE滑落", "⑪償債警報", "⑧現金品質警報"):
            v = r.get(col, "")
            if v: leaks.append(v)
        return " | ".join(leaks)
    h["主要漏洞"] = h.apply(build_leak, axis=1)

    # 統計
    n_eps_lag = (h["⑦EPS落後營收"] != "").sum()
    n_roe_decay = (h["⑪ROE滑落"] != "").sum()
    n_debt = (h["⑪償債警報"] != "").sum()
    n_cash = (h["⑧現金品質警報"] != "").sum()
    n_pen = (h["⑪動態惡化扣分"] < 0).sum()
    print(f"  ⑦ EPS落後營收:       {n_eps_lag} 檔")
    print(f"  ⑪ ROE滑落:           {n_roe_decay} 檔")
    print(f"  ⑪ 償債警報:           {n_debt} 檔")
    print(f"  ⑧ 現金品質警報:       {n_cash} 檔")
    print(f"  ⑪ 動態惡化扣分 ≤ -5:  {n_pen} 檔")

    # ⑪聯動:A級裡有重大惡化的(降級警示)
    a_with_decay = h[(h["評等"] == "A") & (h["⑪動態惡化扣分"] <= -10)]
    print(f"\n⚠️ A級但有⑪重大惡化(該降級觀察)= {len(a_with_decay)} 檔")
    if len(a_with_decay):
        cols = ["代號", "名稱", "ROE%", "ROE5年均%", "含金量", "主要漏洞"]
        cols = [c for c in cols if c in a_with_decay.columns]
        print(a_with_decay[cols].head(15).to_string(index=False))

    # 寫回(保留其他分頁)
    sheets["體檢總表"] = h
    tmp = SRC + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for sh, df in sheets.items():
            df.to_excel(xw, sheet_name=sh, index=False)
    os.replace(tmp, SRC)
    print(f"\n→ 已更新 {SRC}(新增 ⑦/⑪/⑧ 警報欄與動態惡化扣分)")


if __name__ == "__main__":
    main()
