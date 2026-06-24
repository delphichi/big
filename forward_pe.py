# -*- coding: utf-8 -*-
"""
Forward PE / PEG 共用模組(單一真理來源)
=======================================================================
股票買的是未來不是過去:當下 PER 用 TTM EPS,但市場定價未來 EPS。
本模組把「用明年 EPS 修正當下 PE」的邏輯集中一處,供體檢/拐點/財報估值/0050預測共用,
確保各表 forward 口徑一致。

成長率 g:近3年EPS CAGR(主),沒有則退 5年;cap [-30, +60] 防線性外推爆衝。
保守情境 g_保守 = min(g, 月營收YoY) — 月營收領先 EPS,動能轉弱先反映。
循環股豁免:獲利上下震盪,線性外推無意義 → 一律 {},回頭看 PBR。
"""
import pandas as pd

G_FLOOR, G_CAP = -30.0, 60.0


def forward_tag(fpe, peg):
    """Forward PE + PEG 綜合標籤(用未來 EPS 看現價貴不貴)。"""
    if fpe is None:
        return "—"
    if peg is not None and 0 < peg < 1:
        return "🟢成長未反映"          # PEG<1:成長還沒被 price in
    if fpe < 15:
        return "🟢未來便宜"
    if fpe < 22:
        return "🟡未來合理"
    if fpe < 32:
        return "🟠未來偏貴"
    return "🔴未來過熱"


def forward_metrics(close, ttm_eps, per, e3, e5, mo_yoy, cyclical):
    """回傳 forward 估值 dict;循環/資料不足回 {}。
    輸入:收盤、近四季EPS、當下PER、近3年EPS成長%、近5年EPS成長%、月營收YoY%、是否循環。"""
    if cyclical or pd.isna(close) or pd.isna(ttm_eps) or ttm_eps <= 0:
        return {}
    g = e3 if pd.notna(e3) else e5
    if pd.isna(g):
        return {}
    g = max(G_FLOOR, min(G_CAP, float(g)))            # 防線性外推爆衝
    g_cons = g
    if pd.notna(mo_yoy):                               # 保守:月營收動能若更低,以它為準
        g_cons = max(G_FLOOR, min(g, float(mo_yoy)))
    fwd_eps = ttm_eps * (1 + g / 100)
    fwd_eps_c = ttm_eps * (1 + g_cons / 100)
    fpe   = (close / fwd_eps)   if fwd_eps   > 0 else None
    fpe_c = (close / fwd_eps_c) if fwd_eps_c > 0 else None
    peg = (per / g) if (pd.notna(per) and g > 0) else None
    return {"成長率g%": round(g, 1),
            "預估明年EPS": round(fwd_eps, 2),
            "ForwardPE": round(fpe, 1) if fpe else None,
            "ForwardPE保守": round(fpe_c, 1) if fpe_c else None,
            "PEG": round(peg, 2) if peg else None,
            "未來估值": forward_tag(fpe, peg)}
