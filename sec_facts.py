# -*- coding: utf-8 -*-
"""
sec_facts.py — 從 SEC EDGAR companyfacts 算「品質 + 現金含金量/FCF」基本面
================================================================================
定位:不是新 screener,而是一支**共用的基本面資料來源**。
     輸出欄位刻意對齊你 fmp_data.json 的命名(roe_pct / roa_pct / roic_pct /
     gross_margin / interest_cov / fcf_yield_pct ...),好讓現有四支 US screener
     之後可以「換來源不改邏輯」,並能逐欄跟 FMP 對帳。

算什麼(全部用 TTM 流量 + 最新一期存量):
  品質    : ROE / ROA / ROIC / 毛利率 / 淨利率 / 利息覆蓋
  現金含金: OCF(TTM) / FCF(TTM) / FCF Margin(FCF/營收) / OCF÷淨利(現金含金量)
  FCF Yield: 需市值 → 預設 None;呼叫時傳 market_cap 才算(股價來自 yfinance)

為什麼可信:
  - 流量(損益/現金流)用財報原生 fy/fp + 期間天數抓單季,排除 10-Q 的 YTD 累計;
    多數公司缺的 Q4 用「全年 − 前三季」回填(同一 fy,口徑一致)。
  - 存量(資產負債表)是瞬時值,抓最新一期 instant frame。
  - 同一期間多次申報(原始/重編/比較欄)取 filed 最新者(重編優先)。
  - 每檔回傳 missing 清單,算不出來就誠實標 None,不硬湊。

★ 前置:USER_AGENT 在 us_revenue_yoy_scanner.py 設好(SEC 強制帶聯絡 email)。
★ 限制:銀行/保險(無單一 Revenues/GrossProfit 概念)、外國 ADR(交 20-F/IFRS)
        多數欄位會是 None — 這是 SEC 先天盲區,不是錯。金融股請續用 ROA 版 screener。

用法:
  python3 sec_facts.py AAPL NVDA COST            # 印基本面表(含 yfinance 市值算 FCF Yield)
  python3 sec_facts.py --reconcile AAPL NVDA MU  # 跟 fmp_data.json 逐欄對帳
  from sec_facts import sec_facts                  # 程式內呼叫,回傳 dict
"""

import os, sys, json
from datetime import datetime

# 共用 us_revenue_yoy_scanner 的網路 plumbing(USER_AGENT 只設一次、CIK 快取共用)
from us_revenue_yoy_scanner import fetch_facts, load_cik_map, USER_AGENT

# ---------- us-gaap 概念標籤(依優先序合併補洞) ----------
# 流量(duration:損益表 + 現金流量表)
FLOW = {
    "revenue":      ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues",
                     "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"],
    "gross_profit": ["GrossProfit"],
    "cost_of_rev":  ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"],
    "op_income":    ["OperatingIncomeLoss"],
    "net_income":   ["NetIncomeLoss"],
    "interest_exp": ["InterestExpense", "InterestExpenseNonoperating",
                     "InterestExpenseDebt", "InterestAndDebtExpense"],
    "income_tax":   ["IncomeTaxExpenseBenefit"],
    "pretax":       ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                     "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "ocf":          ["NetCashProvidedByUsedInOperatingActivities",
                     "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "capex":        ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
}
# 存量(instant:資產負債表)
STOCK = {
    "assets":            ["Assets"],
    "equity":            ["StockholdersEquity"],
    "debt_noncurrent":   ["LongTermDebtNoncurrent", "LongTermDebt"],
    "debt_current":      ["LongTermDebtCurrent"],
    "debt_short":        ["ShortTermBorrowings", "DebtCurrent", "CommercialPaper"],
}

Q_MIN, Q_MAX = 80, 100      # 單季 ~13 週(含 52/53 週制)
Y_MIN, Y_MAX = 350, 380     # 全年


def _d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()

def _div(a, b):
    return (a / b) if (a is not None and b not in (None, 0)) else None

def _pct(a, b):
    r = _div(a, b)
    return round(r * 100, 2) if r is not None else None


# ---------- 流量:單季 + Q4 回填 + TTM ----------
def _flow_periods(usg, concepts):
    """回傳 (quarters{(fy,q):val}, annuals{fy:val})。多概念依序合併補洞,期間取 filed 最新。"""
    def collect(lo, hi, want_q):
        merged = {}
        for c in concepts:
            node = usg.get(c)
            if not node:
                continue
            local = {}
            for unit, items in node.get("units", {}).items():
                if unit != "USD":
                    continue
                for it in items:
                    s, e, val = it.get("start"), it.get("end"), it.get("val")
                    fp, fy, filed = it.get("fp"), it.get("fy"), it.get("filed", "")
                    if not (s and e) or val is None or fy is None:
                        continue
                    try:
                        days = (_d(e) - _d(s)).days
                    except Exception:
                        continue
                    if not (lo <= days <= hi):
                        continue
                    if want_q:
                        if fp not in ("Q1", "Q2", "Q3", "Q4"):
                            continue
                        key = (int(fy), int(fp[1]))
                    else:
                        key = int(fy)
                    prev = local.get(key)
                    if prev is None or filed > prev[0]:
                        local[key] = (filed, float(val))
            for k, (filed, v) in local.items():
                merged.setdefault(k, v)
        return merged
    return collect(Q_MIN, Q_MAX, True), collect(Y_MIN, Y_MAX, False)

def _with_q4(quarters, annuals):
    """全年 − (Q1+Q2+Q3) 回填缺漏 Q4。回傳 backfilled set。"""
    bf = set()
    for fy in {y for (y, q) in quarters}:
        if (fy, 4) in quarters:
            continue
        if all((fy, q) in quarters for q in (1, 2, 3)) and fy in annuals:
            s3 = quarters[(fy, 1)] + quarters[(fy, 2)] + quarters[(fy, 3)]
            q4 = annuals[fy] - s3
            avg3 = s3 / 3
            if q4 != 0 and (avg3 == 0 or -5 * abs(avg3) <= q4 <= 5 * abs(avg3)):
                quarters[(fy, 4)] = q4          # 流量可為負(虧損季),不強制 > 0
                bf.add((fy, 4))
    return bf

def _ttm(usg, concepts):
    """最近 4 季加總 → TTM。回傳 (ttm值, 最新季key, 用到的季數, 回填季數)。不足 4 季回 None。"""
    q, a = _flow_periods(usg, concepts)
    bf = _with_q4(q, a)
    keys = sorted(q)
    if len(keys) < 4:
        return None, (keys[-1] if keys else None), len(keys), 0
    last4 = keys[-4:]
    val = sum(q[k] for k in last4)
    n_bf = sum(1 for k in last4 if k in bf)
    return val, last4[-1], 4, n_bf


# ---------- 存量:最新一期 instant ----------
def _instant_latest(usg, concepts):
    """資產負債表瞬時值,取 end 最新、同 end 取 filed 最新。回傳 (val, end_date) 或 (None,None)。"""
    best_key, best = None, None       # best_key=(end), tiebreak filed
    for c in concepts:
        node = usg.get(c)
        if not node:
            continue
        for unit, items in node.get("units", {}).items():
            if unit != "USD":
                continue
            for it in items:
                e, val, filed = it.get("end"), it.get("val"), it.get("filed", "")
                # instant 事實沒有 start;若有 start 且非當日,跳過(那是 duration)
                if it.get("start") and it.get("start") != e:
                    continue
                if not e or val is None:
                    continue
                k = (e, filed)
                if best_key is None or k > best_key:
                    best_key, best = k, float(val)
        if best is not None:           # 高優先序概念有值就停,不被低優先序覆蓋
            break
    return (best, best_key[0]) if best is not None else (None, None)


# ---------- 主函數 ----------
def sec_facts(sym, cikmap, facts=None, market_cap=None):
    """
    回傳一檔的品質 + 現金含金量基本面 dict。
    market_cap 有給才算 fcf_yield_pct(SEC 無股價)。
    facts 可預先帶入(避免重複下載);否則自動抓。
    """
    out = {"ticker": sym.upper(), "missing": []}
    cik = cikmap.get(sym.upper()) or cikmap.get(sym.upper().replace(".", "-"))
    if cik is None:
        out["error"] = "查無CIK(非SEC財報/ETF/外國ADR)"
        return out
    if facts is None:
        facts = fetch_facts(cik)
    usg = (facts or {}).get("facts", {}).get("us-gaap", {})
    if not usg:
        out["error"] = "無 us-gaap 財報(可能是外國 IFRS filer)"
        return out

    # --- 流量 TTM ---
    rev,   asof, _,  _    = _ttm(usg, FLOW["revenue"])
    gp,    _,    _,  _    = _ttm(usg, FLOW["gross_profit"])
    cogs,  _,    _,  _    = _ttm(usg, FLOW["cost_of_rev"])
    opinc, _,    _,  _    = _ttm(usg, FLOW["op_income"])
    ni,    asof2, nq, nbf = _ttm(usg, FLOW["net_income"])
    intex, _,    _,  _    = _ttm(usg, FLOW["interest_exp"])
    tax,   _,    _,  _    = _ttm(usg, FLOW["income_tax"])
    pretax,_,    _,  _    = _ttm(usg, FLOW["pretax"])
    ocf,   _,    _,  ocfb = _ttm(usg, FLOW["ocf"])
    capex, _,    _,  _    = _ttm(usg, FLOW["capex"])

    # 毛利:有 GrossProfit 直接用,否則營收 − 銷貨成本
    if gp is None and rev is not None and cogs is not None:
        gp = rev - cogs

    # --- 存量(最新一期) ---
    assets, asof_bs = _instant_latest(usg, STOCK["assets"])
    equity, _       = _instant_latest(usg, STOCK["equity"])
    dnc, _ = _instant_latest(usg, STOCK["debt_noncurrent"])
    dcu, _ = _instant_latest(usg, STOCK["debt_current"])
    dsh, _ = _instant_latest(usg, STOCK["debt_short"])
    debt_parts = [x for x in (dnc, dcu, dsh) if x is not None]
    total_debt = sum(debt_parts) if debt_parts else None

    # --- 衍生:現金流 ---
    fcf = (ocf - capex) if (ocf is not None and capex is not None) else None

    # --- 衍生:ROIC 的 NOPAT(營業利益 × (1−有效稅率)) ---
    tax_rate = _div(tax, pretax)
    if tax_rate is None or not (0 <= tax_rate <= 0.5):
        tax_rate = 0.21                      # 取不到/異常 → 用美國法定稅率近似
    nopat = opinc * (1 - tax_rate) if opinc is not None else None
    invested = (total_debt + equity) if (total_debt is not None and equity is not None) else None

    # --- 組裝(欄名對齊 fmp_data.json) ---
    out.update({
        "asof":            asof2 and f"{asof2[0]}Q{asof2[1]}",   # TTM 截止季(財務年)
        "asof_bs":         asof_bs,                              # 資產負債表日期
        "ttm_quarters":    nq,
        "ttm_q4_backfill": nbf,                                  # TTM 內幾季是回填的
        # 品質
        "roe_pct":         _pct(ni, equity),
        "roa_pct":         _pct(ni, assets),
        "roic_pct":        _pct(nopat, invested),
        "gross_margin":    _pct(gp, rev),
        "net_margin":      _pct(ni, rev),
        "interest_cov":    round(_div(opinc, intex), 2) if (intex and intex > 0 and opinc is not None) else None,
        # 現金含金量 / FCF
        "ocf_ttm_b":       round(ocf / 1e9, 3) if ocf is not None else None,
        "fcf_ttm_b":       round(fcf / 1e9, 3) if fcf is not None else None,
        "fcf_margin":      _pct(fcf, rev),                       # 不需股價,SEC 原生
        "ocf_to_ni":       round(_div(ocf, ni), 2) if (ocf is not None and ni not in (None, 0)) else None,
        "fcf_yield_pct":   _pct(fcf, market_cap),                # 需市值,沒給就 None
        # 原始值(對帳/除錯用)
        "_revenue_ttm_b":  round(rev / 1e9, 3) if rev is not None else None,
        "_ni_ttm_b":       round(ni / 1e9, 3) if ni is not None else None,
        "_equity_b":       round(equity / 1e9, 3) if equity is not None else None,
        "_assets_b":       round(assets / 1e9, 3) if assets is not None else None,
        "_debt_b":         round(total_debt / 1e9, 3) if total_debt is not None else None,
    })
    # missing 診斷
    for k, v in (("revenue", rev), ("net_income", ni), ("equity", equity),
                 ("assets", assets), ("ocf", ocf), ("op_income", opinc)):
        if v is None:
            out["missing"].append(k)
    return out


# ---------- yfinance 市值(選用,算 FCF Yield) ----------
def _market_cap(sym):
    try:
        import yfinance as yf
        fi = yf.Ticker(sym).fast_info
        mc = fi.get("market_cap") if hasattr(fi, "get") else getattr(fi, "market_cap", None)
        return float(mc) if mc else None
    except Exception:
        return None


# ---------- CLI ----------
def _guard_ua():
    if "your_email@example.com" in USER_AGENT or USER_AGENT.startswith("ChangeMe"):
        print("⚠ 請先到 us_revenue_yoy_scanner.py 把 USER_AGENT 改成『你的名字 你的email』再執行。")
        sys.exit(1)

def cmd_table(syms):
    cikmap = load_cik_map()
    rows = []
    for s in syms:
        mc = _market_cap(s)
        rows.append(sec_facts(s, cikmap, market_cap=mc))
    hdr = ["ticker", "asof", "roe_pct", "roa_pct", "roic_pct", "gross_margin",
           "net_margin", "interest_cov", "fcf_margin", "ocf_to_ni", "fcf_yield_pct", "ttm_q4_backfill"]
    print("  ".join(f"{h:>12}" for h in hdr))
    for r in rows:
        if r.get("error"):
            print(f"{r['ticker']:>12}  {r['error']}")
            continue
        print("  ".join(f"{str(r.get(h, '')):>12}" for h in hdr))
        if r.get("missing"):
            print(f"{'':>12}  (missing: {', '.join(r['missing'])})")

def cmd_reconcile(syms):
    """跟 fmp_data.json 逐欄對帳,看 SEC 算的跟 FMP 差多少。"""
    # 找 fmp_data.json(沿用你 reversal-screener 的位置)
    cands = ["fmp_data.json", "reversal-screener/fmp_data.json",
             os.path.expanduser("~/outputs/reversal-screener/fmp_data.json")]
    fmp_path = next((p for p in cands if os.path.exists(p)), None)
    if not fmp_path:
        print("找不到 fmp_data.json(試過:%s)" % ", ".join(cands)); return
    fmp = json.load(open(fmp_path, encoding="utf-8"))
    print(f"對帳基準:{fmp_path}\n")
    cikmap = load_cik_map()
    pairs = [("roe_pct", "roe_pct"), ("roa_pct", "roa_pct"), ("roic_pct", "roic_pct"),
             ("gross_margin", "gross_margin"), ("net_margin", "net_margin")]
    for s in syms:
        sec = sec_facts(s, cikmap)
        f = fmp.get(s.upper(), {})
        print(f"=== {s.upper()}  (SEC asof {sec.get('asof')}) ===")
        if sec.get("error"):
            print("  SEC:", sec["error"]); continue
        for sk, fk in pairs:
            sv, fv = sec.get(sk), f.get(fk)
            if sv is None and fv is None:
                continue
            diff = (f"{sv - fv:+.2f}" if (sv is not None and fv is not None) else "—")
            flag = ""
            if sv is not None and fv not in (None, 0):
                if abs(sv - fv) / abs(fv) > 0.10:
                    flag = "  ⚠差>10%"
            print(f"  {sk:14} SEC={str(sv):>8}  FMP={str(fv):>8}  Δ={diff}{flag}")
        if sec.get("missing"):
            print(f"  (SEC missing: {', '.join(sec['missing'])})")
        print()


def main():
    _guard_ua()
    args = sys.argv[1:]
    if not args:
        print("用法:python3 sec_facts.py [--reconcile] TICKER [TICKER ...]"); return
    if args[0] == "--reconcile":
        cmd_reconcile(args[1:])
    else:
        cmd_table(args)


if __name__ == "__main__":
    main()
