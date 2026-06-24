# -*- coding: utf-8 -*-
"""
財報停損價共用模組 — 用估值錨自動算每檔的「估值支撐被跌破」價格
=======================================================================
取代「現價 × 0.85」這種隨便拍的停損,改用財報數據算真實支撐位。

三種估值錨,看股票類型選用:
  PE 法 :成長股(EPS 穩成長)→ EPS × PE5年均 × 0.85
  PBR法 :資產/循環股(獲利波動)→ 每股淨值 × PBR5年均 × 0.85
  殖利率法:高息/定存股 → 年股利 ÷ (殖利率5年均高位 × 1.15)

選錨規則(基於體檢/財報資料):
  - 循環股(獲利震盪) → PBR
  - 評等A/B + EPS 連年成長 + PEG 可算 → PE
  - 殖利率 > 5% 且 EPS 穩定(波動小)→ 殖利率法
  - 預設 → PE,並輸出三種供對照
"""
import pandas as pd

SAFETY_MARGIN = 0.85   # 歷史均值 × 0.85 = 合理區間下緣


def calc_anchors(close, eps_ttm, pe5y, pbr_cur, pbr5y, dy_cur, dy5y):
    """回傳三種錨的停損價 dict;算不出回 None。"""
    out = {}
    if pd.notna(eps_ttm) and pd.notna(pe5y) and eps_ttm > 0 and pe5y > 0:
        out["PE法"] = round(eps_ttm * pe5y * SAFETY_MARGIN, 1)
    if pd.notna(close) and pd.notna(pbr_cur) and pd.notna(pbr5y) and pbr_cur > 0 and pbr5y > 0:
        bvps = close / pbr_cur
        out["PBR法"] = round(bvps * pbr5y * SAFETY_MARGIN, 1)
    if pd.notna(close) and pd.notna(dy_cur) and pd.notna(dy5y) and dy_cur > 0 and dy5y > 0:
        # 反推年股利 = 現價 × 殖利率現/100;殖利率高位 ≈ 5年均 × 1.15
        dividend = close * dy_cur / 100
        dy_high = dy5y * 1.15
        if dy_high > 0:
            out["殖利率法"] = round(dividend / (dy_high / 100), 1)
    return out


def pick_primary_anchor(anchors, cyclical, dy5y, eps_growth_3y):
    """選主錨。回傳 (主錨名稱, 主錨停損價, 理由)。"""
    if not anchors:
        return None, None, "資料不足"
    # 循環股:看 PBR
    if cyclical and "PBR法" in anchors:
        return "PBR法", anchors["PBR法"], "循環股(獲利震盪),看淨值支撐"
    # 高息穩定股(殖利率>5%、EPS低波動):用殖利率法
    if pd.notna(dy5y) and dy5y >= 5 and pd.notna(eps_growth_3y) and abs(eps_growth_3y) < 8:
        if "殖利率法" in anchors:
            return "殖利率法", anchors["殖利率法"], "高息穩定股,殖利率高位有買盤"
    # 預設:成長股看 PE
    if "PE法" in anchors:
        return "PE法", anchors["PE法"], "成長股,看獲利倍數支撐"
    # 退而求其次
    first = list(anchors.keys())[0]
    return first, anchors[first], "預設"


def stop_recommendation(user_stop, primary_stop, close):
    """比對使用者手填停損 vs 財報停損 → 建議。"""
    if primary_stop is None or user_stop is None:
        return ""
    diff_pct = (user_stop - primary_stop) / primary_stop * 100
    if abs(diff_pct) < 5:
        return "✅ 手填停損與財報支撐一致"
    if user_stop > primary_stop:
        return f"⚠️ 手填停損 {user_stop} 太緊(財報支撐在 {primary_stop:.0f},高 {diff_pct:.0f}%)→ 易被洗出"
    else:
        return f"⚠️ 手填停損 {user_stop} 太鬆(財報支撐在 {primary_stop:.0f},低 {-diff_pct:.0f}%)→ 多賠"


def analyze(sid, row_val, row_hist, row_health, user_stop):
    """完整分析:回 dict 包含三錨+主錨+建議。
    row_val: 財報估值比較那列;row_hist: 相對歷史水位那列;row_health: 體檢總表那列。"""
    close = pd.to_numeric(row_val.get("收盤"), errors="coerce")
    eps   = pd.to_numeric(row_val.get("近四季EPS"), errors="coerce")
    pe5y  = pd.to_numeric(row_hist.get("PER5年均"), errors="coerce") if row_hist is not None else None
    pbr_c = pd.to_numeric(row_hist.get("PBR現"), errors="coerce") if row_hist is not None else None
    pbr5y = pd.to_numeric(row_hist.get("PBR5年均"), errors="coerce") if row_hist is not None else None
    dy_c  = pd.to_numeric(row_hist.get("殖利率現"), errors="coerce") if row_hist is not None else None
    dy5y  = pd.to_numeric(row_hist.get("殖利率5年均"), errors="coerce") if row_hist is not None else None
    cyc   = "循環" in str(row_health.get("循環股", "")) if row_health is not None else False
    e3    = pd.to_numeric(row_health.get("EPS近3y%"), errors="coerce") if row_health is not None else None

    anchors = calc_anchors(close, eps, pe5y, pbr_c, pbr5y, dy_c, dy5y)
    name, price, why = pick_primary_anchor(anchors, cyc, dy5y, e3)
    rec = stop_recommendation(user_stop, price, close)
    return {
        "現價": close,
        "PE法停損": anchors.get("PE法"),
        "PBR法停損": anchors.get("PBR法"),
        "殖利率法停損": anchors.get("殖利率法"),
        "建議主錨": name,
        "建議停損價": price,
        "選錨理由": why,
        "手填停損": user_stop,
        "停損建議": rec,
    }
