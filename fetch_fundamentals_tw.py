# -*- coding: utf-8 -*-
"""
台股 財報 + 估值 表 (TW Fundamentals & Valuation)
=======================================================================
名單 = 月營收年增掃描中表現最好的一批標的(共 247 檔,含金融股,見 PICKS)。
不滿 5 年資料者(新上市/興櫃轉上市)會自動跳過 5 年趨勢(CAGR / 5年均 ROE / 淨利率留白),
仍照常輸出近四季績效與估值(PER/PBR/殖利率),不影響整體跑批。
一次抓最近 ~5 年的三大表 + 股價 + 每日 PER/PBR/殖利率,算出 5 年趨勢、近四季經營績效、
估值與「相對歷史水位」,輸出可排序的跨檔比較表 +「相對歷史水位」+「逐年營收/EPS」。
股名由 FinMind taiwan_stock_info() 取官方名稱(避免人工標錯);金融股以「🏦」標記。

★ PER 一律「自算」,不用 FinMind 的 per_pbr「PER」欄位 ★
  FinMind 的 per_pbr「PER」基準 EPS 不一致、與 Goodinfo/財報狗對不上(例:奇鋐顯示 12 但
  實際 ~47、台達電 19 但實際 ~79)。改成與 Goodinfo 本益比河流圖同一套:
      PER = 收盤價 ÷ 近四季EPS,近四季EPS = 最近 4 個單季 EPS 加總(FinMind 財報已驗證準確)。
  歷史 PER 序列以 merge_asof(收盤價對近四季EPS)逐日重算,並加上財報公布落後天數
  (Q4 年報 +90 天、其餘 +45 天)避免未卜先知;PER 位階% 由此自算序列計算。
  PBR 與 殖利率 仍取 per_pbr(價格基礎,可靠)。

資料來源:FinMind(taiwan_stock_financial_statement / balance_sheet /
          cash_flows_statement / daily / per_pbr / month_revenue)。一次呼叫即回傳整段歷史,
          故抓 5 年與抓 3 年的 API 次數相同(每檔 6 次)。
輸出   :data/台股財報估值.xlsx

每檔算出:
  5 年趨勢 → 5年營收CAGR%、5年平均淨利率%、5年平均ROE%(只取季數=4 的完整年)
  近四季   → 毛利率 / 營益率 / 淨利率、近四季EPS、近四季ROE
  財務結構 → 負債比、流動比(最新季)
  現金流照妖鏡 → 獲利含金量(近四季營業現金流 ÷ 近四季淨利)、近四季自由現金流(億)
  估值     → 收盤、PER(自算)、PE位階%、PBR、殖利率%
  成長     → 最新月營收年增%

★ 大量抓取務必設環境變數 FINMIND_TOKEN(免費約 300 次/hr、設 token 約 600 次/hr);
  247 檔 × 6 dataset ≈ 1482 次呼叫 → 跨多個整點,務必設 token + MAX_RUNTIME_MIN 控時。
  斷點續跑:每檔算完即存 data/_tw_val_cache/{代號}.json,並每 10 檔重建一次 Excel;
  撞額度會「睡到整點再續」,被取消/逾時也不丟進度——再跑一次會自動跳過已完成、接著抓。
"""

import os, time, json
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from forward_pe import forward_metrics      # 未來估值(Forward PE/PEG)單一真理來源

# ---------- 設定 ----------
TOKEN      = os.environ.get("FINMIND_TOKEN", "")
START_DATE = "2020-01-01"                 # 取 ~5 完整年(2021–2025)+ 年增基期
OUTPUT     = "data/台股財報估值.xlsx"
RATE_SLEEP = 0.4                          # 每檔間隔(降低撞限流機率)
WRITE_DETAIL = False                      # 逐季明細會產生很多分頁,預設關;要看單檔細節再開
MAX_RUNTIME_MIN = 30                       # 整輪上限(分):到時存檔退出,靠快取下輪續抓 → 控 Actions 計費
RATE_WAIT_CAP   = 90                       # 撞額度時最多睡幾秒(不再睡到整點,避免燒分鐘)
MAX_RATE_RETRY  = 2                        # 撞額度短重試次數,仍失敗就跳過此檔(下輪靠快取補)

# 名單 = 月營收年增掃描中表現最好的一批標的 = 98 檔(含金融股,口徑不適用者輸出會標 🏦)。
# 只放代號,股名在 CI 由 FinMind taiwan_stock_info() 取官方名稱(避免人工標錯,如 2947 振宇五金)。
PICKS = [
    "6721","6741","4953","3293","3587","2753","6690","2947","6223","2453",
    "3017","3004","3583","6279","2890","6803","2755","4129","6446","2912",
    "2345","2376","6752","6733","2640","2480","2308","2752","2937","2402",
    "5278","9917","8284","6612","3265","4114","3653","5903","9911","3130",
    "1513","4116","2812","2330","1736","5904","2880","2850","2887","3036",
    "3162","3029","5493","2059","2368","8462","2383","5607","5274","5312",
    "6112","8210","6449","3687","6257","6683","3551","6776","7556","6138",
    "6525","3044",
    # 追加 26 檔
    "1519","2353","2357","2363","2645","2884","2889","2891","3019","3324",
    "3630","3702","3715","4772","4904","5203","5269","6005","6274","6469",
    "6712","6739","6787","6788","6791","8016",
    # 追加餐飲同業
    "1268","2727",
    # 追加散熱/液冷同業
    "8996","3013",
    # 追加半導體設備/IC
    "1560","2360","2379",
    # 追加光學鏡頭同業
    "3008","3406","3362","6209","3504",
    # 追加 CPO/矽光子/光通訊/先進封裝/測試
    "3081","2455","4991","3363","4979","6442","3163","4977","3234","3711",
    "6451","3450","6515",
    # 追加 茂順(觀察清單帶入)
    "9942",
    # 追加 金融/公用/石化
    "2882","2881","2851","2885","2883","2855","8926","6505",
    # 追加 重電/工具機
    "1504","2049",
    # 追加 環保/工程
    "8341","8422","9933",
    # 追加 0050候補名單
    "2313","3481","3189",
    # 追加 觀察批(EPS成長篩選)
    "4174","4746","1784","6589","8299","2408","2344","2337","6770","5289",
    "2329","3260","2451","8150","2395","2065","4967","8271","5227","6863",
    "6441","2327","4583","2915","3105","1723","3023","3016","6488","6182",
    "8231","6435","2634","1503","1514","8027","3673","8064","3131","3455",
    "3055","3535","4760","4768","3037","2467","1802","4416","2412","3045",
    "2404","4749","9931","9918","9908","3026",
    # 追加 觀察批#2
    "6949","2371","6239","2425","8719","7703","7730","3030","7825","7828",
    "1595","4542","3149","4728","2421","3005","4506","2535","1215","1232",
    "5434","1773",
    # 追加 2301 光寶科 / 8081 致新
    "2301","8081",
    # 0050 2026Q2 納入4檔(回測拐點是否提前出現)
    "8046","3443","3665","4958",
    # 0050 + 富櫃50 blind spot 補抓(依 stock_etf_inclusion_predict 輸出補)
    "3661","3533","3034","1590","2207","2356","6415","6139","1605",
    "3227","5371","8086","5351","4123","6613","3357","5439","5425","6640",
    "5347","3529","8069","6147",
    # 偵錯批:營建/金融/電子/半導體設備材料(加入快照守門,監看判讀漂移)
    "2515","3703","6177","2509","1438","5213","2886","5880","4938","5309",
    "5292","8176","6944","7818","2013","6486","1521","6196","6826","4755",
    "4722","2422","3680","1764","3532","6937",
]

# 金融股(銀行/金控/保險/券商):營收/利潤率口徑不適用,輸出會標 🏦(估值 PER/PBR/殖利率仍有效)。
FINANCIALS = {"2890","2884","2880","2887","2891","2812","2836","6005","2850","2889",
               "2882","2881","2851","2885","2883","2855"}


# ---------- FinMind ----------
def make_loader():
    from FinMind.data import DataLoader
    dl = DataLoader()
    if TOKEN:
        try:
            dl.login_by_token(api_token=TOKEN)
        except Exception as e:
            print("token 登入失敗(改用免費額度):", e)
    return dl

def load_names(dl):
    """從 FinMind 取『代號→官方股名』對照(避免人工標錯)。失敗則回空 dict。"""
    try:
        info = dl.taiwan_stock_info()
        return {str(r["stock_id"]): str(r["stock_name"]) for _, r in info.iterrows()}
    except Exception as e:
        print("取股名對照失敗(改用代號):", e)
        return {}

def _is_rate_limit(e):
    msg = str(e).lower()
    return any(k in msg for k in ("limit", "402", "429", "too many", "exceed", "request"))

def seconds_to_next_hour(buffer=45):
    """距下一個整點還有幾秒(FinMind 額度每小時重置),多加 buffer 秒保險。"""
    now = datetime.now()
    nxt = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    return max(5, int((nxt - now).total_seconds()) + buffer)

def get_per(dl, sid, start):
    """每日 PER/PBR/殖利率(只取 PBR 與殖利率;PER 自算)。先試 DataLoader,失敗退回原生 REST。"""
    try:
        return dl.taiwan_stock_per_pbr(stock_id=sid, start_date=start)
    except Exception:
        import requests
        h = {"Authorization": f"Bearer {TOKEN}"} if TOKEN else {}
        r = requests.get("https://api.finmindtrade.com/api/v4/data",
                         params={"dataset": "TaiwanStockPER", "data_id": sid, "start_date": start},
                         headers=h, timeout=20)
        return pd.DataFrame(r.json().get("data", []))

def fetch_one(dl, sid, start):
    out = {
        "損益表":    dl.taiwan_stock_financial_statement(stock_id=sid, start_date=start),
        "資產負債表": dl.taiwan_stock_balance_sheet(stock_id=sid, start_date=start),
        "現金流量表": dl.taiwan_stock_cash_flows_statement(stock_id=sid, start_date=start),
        "股價":      dl.taiwan_stock_daily(stock_id=sid, start_date=start),
        "月營收":    dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start),
        "PER":      get_per(dl, sid, start),
    }
    return out


# ---------- 工具 ----------
def pivot(df):
    if df is None or df.empty or "type" not in df.columns:
        return pd.DataFrame()
    return df.pivot_table(index="date", columns="type", values="value", aggfunc="first").sort_index()

def pick(piv, *names):
    for n in names:
        if n in piv.columns:
            return piv[n]
    return pd.Series(index=piv.index, dtype="float64")

def decum(s):
    """台股『現金流量表』是 YTD 累計(Q2=半年、Q3=前三季、Q4=全年);
    轉成單季:同一年『本期 − 上一期』,每年首季維持原值。損益表已是單季,不需處理。"""
    if s is None or len(s) == 0:
        return s
    s = s.sort_index()
    out, prev_y, prev_v = {}, None, None
    for d, v in s.items():
        y = str(d)[:4]
        if pd.isna(v):
            out[d] = v
            continue
        out[d] = (v - prev_v) if (y == prev_y and prev_v is not None) else v
        prev_y, prev_v = y, v
    return pd.Series(out)


# ---------- 自算 PER(收盤價 ÷ 近四季EPS,含公布落後)----------
def ttm_eps(inc_df):
    """單季 EPS → 近四季EPS 序列(index=季底日,加上公布落後天數當『生效日』,避免未卜先知)。"""
    piv = pivot(inc_df)
    if piv.empty or "EPS" not in piv.columns:
        return None
    eps = pd.to_numeric(piv["EPS"], errors="coerce").dropna()
    if len(eps) < 4:
        return None
    ttm = eps.rolling(4).sum().dropna()
    rows = []
    for d, v in ttm.items():
        qend = pd.to_datetime(d)
        lag = 90 if qend.month == 12 else 45        # Q4(年報)約90天、其餘約45天才公布
        rows.append((qend + timedelta(days=lag), float(v)))
    return pd.DataFrame(rows, columns=["生效日", "近四季EPS"]).sort_values("生效日")

def per_series(price_df, inc_df):
    """逐日自算 PER = 收盤價 ÷ 近四季EPS(merge_asof 向後對齊,只取 EPS>0 的交易日)。
    回傳含 date / close / 近四季EPS / PER 的 DataFrame;資料不足回 None。"""
    tt = ttm_eps(inc_df)
    if tt is None or tt.empty:
        return None
    if price_df is None or price_df.empty or "close" not in price_df.columns:
        return None
    p = price_df[["date", "close"]].copy()
    p["close"] = pd.to_numeric(p["close"], errors="coerce")
    p = p.dropna().sort_values("date")
    if p.empty:
        return None
    p["生效日"] = pd.to_datetime(p["date"])
    m = pd.merge_asof(p.sort_values("生效日"), tt, on="生效日", direction="backward")
    m = m[(m["近四季EPS"] > 0)].copy()
    if m.empty:
        return None
    m["PER"] = m["close"] / m["近四季EPS"]
    return m


# ---------- 逐季經營績效 ----------
def performance(raw):
    inc, bal, cf = pivot(raw["損益表"]), pivot(raw["資產負債表"]), pivot(raw["現金流量表"])
    if inc.empty:
        return pd.DataFrame()

    rev  = pick(inc, "Revenue")
    gp   = pick(inc, "GrossProfit")
    op   = pick(inc, "OperatingIncome")
    ni   = pick(inc, "IncomeAfterTaxes", "ProfitAfterTax", "NetIncome")
    eps  = pick(inc, "EPS")

    ta   = pick(bal, "TotalAssets", "Total_Assets")
    tl   = pick(bal, "TotalLiabilities", "Liabilities", "Total_Liabilities")
    eq   = pick(bal, "Equity", "TotalEquity", "EquityAttributableToOwnersOfParent")
    ca   = pick(bal, "CurrentAssets")
    cl   = pick(bal, "CurrentLiabilities")
    cap  = pick(bal, "CommonStocks", "CommonStock", "CapitalStock", "Capital", "ShareCapital")
    inv  = pick(bal, "Inventories", "Inventory")           # 存貨:用於計算存貨年增vs營收年增是否背離

    # 現金流量表為 YTD 累計 → 先轉單季,之後加總近四季才正確
    ocf  = decum(pick(cf, "CashFlowsFromOperatingActivities",
                          "NetCashFlowsFromOperatingActivities",
                          "CashProvidedByOperatingActivities"))
    capex = decum(pick(cf, "PropertyAndPlantAndEquipment",
                          "AcquisitionOfPropertyPlantAndEquipment",
                          "PaymentsToAcquirePropertyPlantAndEquipment"))

    m = pd.DataFrame(index=inc.index)
    m["營收(億)"]   = (rev / 1e8).round(1)
    m["_毛利"]      = gp
    m["_營益"]      = op
    m["_淨利"]      = ni
    m["毛利率%"]    = (gp / rev * 100).round(2)
    m["營益率%"]    = (op / rev * 100).round(2)
    m["淨利率%"]    = (ni / rev * 100).round(2)
    m["EPS"]       = eps.round(2)
    m["_權益"]      = eq
    m["負債比%"]    = (tl / ta * 100).round(1)
    m["流動比%"]    = (ca / cl * 100).round(0)
    m["_OCF"]       = ocf
    m["營業現金流(億)"] = (ocf / 1e8).round(1)
    m["自由現金流(億)"] = ((ocf + capex) / 1e8).round(1)   # capex 在現金流量表多為負值
    m["_股本"]       = cap                                # 台股股本(元),÷10 為流通股數
    m["_存貨"]       = inv                                # 存貨原值(元),供算存貨年增與營收背離
    return m


def revenue_yoy(raw):
    rv = raw["月營收"]
    if rv is None or rv.empty:
        return None
    rv = rv.sort_values("date").reset_index(drop=True)
    yoy = rv["revenue"].pct_change(12) * 100
    return round(float(yoy.iloc[-1]), 1) if len(yoy) and pd.notna(yoy.iloc[-1]) else None


# ---------- 逐年彙整(供 5 年比較)----------
def yearly(perf):
    """把逐季 perf 彙整成『逐年』:年度營收(億)/EPS 加總、年末權益、淨利、三率年均、季數。"""
    yd = {}
    for d, r in perf.iterrows():
        y = str(d)[:4]
        o = yd.setdefault(y, {"rev": 0.0, "eps": 0.0, "ni": 0.0, "eq": None,
                              "gm": [], "om": [], "nm": [], "qn": 0})
        o["qn"] += 1
        if pd.notna(r["營收(億)"]): o["rev"] += r["營收(億)"]
        if pd.notna(r["EPS"]):     o["eps"] += r["EPS"]
        if pd.notna(r["_淨利"]):    o["ni"]  += r["_淨利"]
        if pd.notna(r["_權益"]):    o["eq"]   = r["_權益"]          # 年內最後一筆 ≈ 年末權益
        for k, col in (("gm", "毛利率%"), ("om", "營益率%"), ("nm", "淨利率%")):
            if pd.notna(r[col]): o[k].append(r[col])
    return yd


# ---------- 相對歷史水位(現值 vs 5年均 vs 位階)----------
def _ttm_series(perf):
    """逐季滾動 TTM 三率與 ROE(近四季加總,口徑一致),供算歷史分布。"""
    rows = []
    for i in range(3, len(perf)):
        win = perf.iloc[i - 3:i + 1]
        rev = win["營收(億)"].sum() * 1e8
        if not rev or pd.isna(rev):
            continue
        ni = win["_淨利"].sum()
        eq = win["_權益"].dropna()
        rows.append({
            "gm": win["_毛利"].sum() / rev * 100,
            "om": win["_營益"].sum() / rev * 100,
            "nm": ni / rev * 100,
            "roe": (ni / eq.iloc[-1] * 100) if (len(eq) and eq.iloc[-1]) else None,
        })
    return rows

def hist_levels(sid, name, perf, per_df, per_ser):
    """回傳一檔的『現值 / 5年均 / 位階%』。位階% = 目前值 ≤ 之歷史比例(0–100)。
    PER 位階用『自算 PER 序列』(per_ser);PBR/殖利率用 per_pbr(價格基礎,可靠)。"""
    out = {"代號": sid, "名稱": name}
    def lvl(vals, cur, prefix, dec=1):
        v = pd.Series([x for x in vals if x is not None and pd.notna(x)], dtype="float64")
        if len(v) and cur is not None and pd.notna(cur):
            out[f"{prefix}現"]   = round(float(cur), dec)
            out[f"{prefix}5年均"] = round(float(v.mean()), dec)
            out[f"{prefix}位階%"] = round(float((v <= cur).mean() * 100))
    ttm = _ttm_series(perf)
    for key, label in (("gm", "毛利率"), ("om", "營益率"), ("nm", "淨利率"), ("roe", "ROE")):
        series = [r[key] for r in ttm]
        cur = series[-1] if series else None
        lvl(series, cur, label)
    # PER:自算序列(順便存 P10/P25/P50 三層 PE 分位點,供「合理價鬧鐘」用)
    if per_ser is not None and not per_ser.empty:
        s = per_ser["PER"].replace([float("inf")], pd.NA).dropna()
        s = s[s > 0]
        if len(s):
            lvl(list(s), float(s.iloc[-1]), "PER", 2)
            out["PE_P10"] = round(float(s.quantile(0.10)), 2)   # 歷史低檔(深度買點)
            out["PE_P25"] = round(float(s.quantile(0.25)), 2)   # 偏便宜
            out["PE_P50"] = round(float(s.quantile(0.50)), 2)   # 中位合理
    # PBR / 殖利率:per_pbr
    if per_df is not None and not per_df.empty:
        p = per_df.sort_values("date")
        for col, label, dec in (("PBR", "PBR", 2), ("dividend_yield", "殖利率", 2)):
            s = pd.to_numeric(p.get(col), errors="coerce").dropna()
            if label == "PBR":
                s = s[s > 0]
            if len(s):
                lvl(list(s), float(s.iloc[-1]), label, dec)
    return out


# ---------- 跨檔比較(一檔一列)----------
def summary_row(sid, name, raw):
    row = {"代號": sid, "名稱": name}
    rev_year, eps_year = {}, {}                       # 供「逐年」跨檔對照表
    perf = performance(raw)
    if not perf.empty:
        # ── 5 年趨勢(只取季數=4 的完整年)──
        yd = yearly(perf)
        full = sorted(y for y, o in yd.items() if o["qn"] >= 4)[-5:]   # 最近 5 個完整年
        rev_year = {y: round(yd[y]["rev"], 1) for y in full}
        eps_year = {y: round(yd[y]["eps"], 2) for y in full}
        if len(full) >= 2 and yd[full[0]]["rev"] > 0:
            n = len(full) - 1
            row["5年營收CAGR%"] = round(((yd[full[-1]]["rev"] / yd[full[0]]["rev"]) ** (1 / n) - 1) * 100, 1)
        roe_y = [yd[y]["ni"] / yd[y]["eq"] * 100 for y in full if yd[y]["eq"]]
        nm_y  = [float(np.mean(yd[y]["nm"])) for y in full if yd[y]["nm"]]
        if roe_y: row["5年平均ROE%"]   = round(float(np.mean(roe_y)), 1)
        if nm_y:  row["5年平均淨利率%"] = round(float(np.mean(nm_y)), 1)
        last4 = perf.tail(4)
        # 近四季平均三率(平滑單季波動)
        row["毛利率%"]   = round(float(last4["毛利率%"].mean(skipna=True)), 1)
        row["營益率%"]   = round(float(last4["營益率%"].mean(skipna=True)), 1)
        row["淨利率%"]   = round(float(last4["淨利率%"].mean(skipna=True)), 1)
        # 近四季 EPS、ROE(近四季淨利 ÷ 最新季權益)
        eps4 = last4["EPS"].dropna()
        row["近四季EPS"] = round(float(eps4.sum()), 2) if len(eps4) else None
        ni4  = last4["_淨利"].dropna().sum()
        eqL  = perf["_權益"].dropna()
        if len(eqL) and eqL.iloc[-1]:
            row["近四季ROE%"] = round(ni4 / eqL.iloc[-1] * 100, 1)
        # 財務結構(最新季)
        row["負債比%"] = perf["負債比%"].dropna().iloc[-1] if perf["負債比%"].notna().any() else None
        row["流動比%"] = perf["流動比%"].dropna().iloc[-1] if perf["流動比%"].notna().any() else None
        # 存貨年增%:最新季存貨 vs 去年同季,>40%且同期營收年增<10% = 需求軟警訊(八方案例)
        inv_ser = perf["_存貨"].dropna() if "_存貨" in perf.columns else pd.Series(dtype=float)
        if len(inv_ser) >= 5 and inv_ser.iloc[-5] > 0:
            row["存貨年增%"] = round((inv_ser.iloc[-1] / inv_ser.iloc[-5] - 1) * 100, 1)
        # 現金流照妖鏡(近四季營業現金流 ÷ 近四季淨利)
        ocf4 = last4["_OCF"].dropna().sum()
        if ni4:
            row["獲利含金量"] = round(ocf4 / ni4, 2)
        row["近四季自由現金流(億)"] = round(float(last4["自由現金流(億)"].sum(skipna=True)), 1)
        row["最新季"] = str(perf.index[-1])

    # 估值:PER 自算(收盤 ÷ 近四季EPS);PBR / 殖利率 取 per_pbr
    ps = per_series(raw.get("股價"), raw.get("損益表"))
    if ps is not None and not ps.empty:
        s = ps["PER"].replace([float("inf")], pd.NA).dropna()
        s = s[s > 0]
        if len(s):
            row["收盤"]      = round(float(ps["close"].iloc[-1]), 1)
            row["PER(自算)"] = round(float(s.iloc[-1]), 2)
            row["PE位階%"]   = round(float((s <= s.iloc[-1]).mean() * 100))
    # 市值(億):股本 ÷ 10 = 流通股數;× 收盤 = 市值;÷ 1e8 換算億
    if not perf.empty and "_股本" in perf.columns and row.get("收盤") is not None:
        cap_ser = perf["_股本"].dropna()
        if len(cap_ser):
            shares = float(cap_ser.iloc[-1]) / 10.0
            row["市值(億)"] = round(shares * float(row["收盤"]) / 1e8, 1)
    per = raw.get("PER")
    if per is not None and not per.empty:
        p = per.sort_values("date").iloc[-1]
        row["PBR"]    = p.get("PBR")
        row["殖利率%"] = p.get("dividend_yield")
    # 成長:最新月營收年增
    row["最新月營收年增%"] = revenue_yoy(raw)
    # ROIC 真實版:近四季營業利益 × 稅後 / 最新季權益
    # (簡化版投入資本只用權益;真實 ROIC 還要加長債,但台股財報長債需另外拆,先用權益保守估)
    # 用途:分辨「資本配置有效」vs「越投越爛」— 例:統一超這種大資本支出股
    if not perf.empty and "_營益" in perf.columns and "_權益" in perf.columns:
        op4 = perf["_營益"].dropna().tail(4).sum()
        eqL = perf["_權益"].dropna()
        if len(eqL) and eqL.iloc[-1] and op4:
            # NOPAT 假設稅率 20%(台股營所稅);除以期末權益
            roic = (op4 * 0.8) / float(eqL.iloc[-1]) * 100
            row["ROIC%(真實)"] = round(float(roic), 1)
    # 未來估值(Forward PE/PEG)— 與體檢/拐點同口徑(forward_pe 共用模組);循環股自動豁免
    ev = [eps_year[y] for y in sorted(eps_year)] if eps_year else []
    if len(ev) >= 2:
        e5 = ((ev[-1]/ev[0])**(1/(len(ev)-1))-1)*100 if ev[0] > 0 and ev[-1] > 0 else np.nan
        e3 = ((ev[-1]/ev[-3])**0.5-1)*100 if len(ev) >= 3 and ev[-3] > 0 and ev[-1] > 0 else np.nan
        cyc = (min(ev) <= 0) or any(ev[i] < ev[i-1]*0.8 for i in range(1, len(ev)))
        fwd = forward_metrics(row.get("收盤"), row.get("近四季EPS"), row.get("PER(自算)"),
                              e3, e5, row.get("最新月營收年增%"), cyc)
        for k, vv in fwd.items():
            row[k] = vv
    hist = hist_levels(sid, name, perf, raw.get("PER"), ps) if not perf.empty else {"代號": sid, "名稱": name}
    # 合理價鬧鐘:歷史 PE 分位 × forward EPS = 三層觸發價(取代「PE 位階」相對排名,給絕對價位)
    # forward EPS 優先用 預估明年EPS,缺則退回 近四季EPS;循環股(已被 forward_pe 豁免)用近四季
    fwd_eps = row.get("預估明年EPS") or row.get("近四季EPS")
    if pd.notna(fwd_eps) and fwd_eps > 0:
        for q, label in (("PE_P50", "合理價"), ("PE_P25", "偏便宜價"), ("PE_P10", "深度買點價")):
            pe_q = hist.get(q)
            if pe_q is not None and pd.notna(pe_q):
                row[label] = round(float(fwd_eps) * float(pe_q), 1)
    return row, perf, rev_year, eps_year, hist


# ---------- 斷點續存(每檔算完即存,可續跑;被取消也不丟進度)----------
CACHE_DIR = "data/_tw_val_cache"

def _cache_path(sid):
    return os.path.join(CACHE_DIR, f"{sid}.json")

def save_cache(sid, obj):
    os.makedirs(CACHE_DIR, exist_ok=True)
    tmp = _cache_path(sid) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, default=float)   # numpy 數值轉 float
    os.replace(tmp, _cache_path(sid))                          # 原子寫入,避免半截檔

def load_cache(sid):
    p = _cache_path(sid)
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


# ---------- 由快取組出 Excel(每次都用『目前所有已完成』重建,故隨時有最新進度檔)----------
def build_output(namemap):
    rows, hists, rev_years, eps_years, q_gms = [], [], {}, {}, {}
    done = 0
    for sid in PICKS:
        c = load_cache(sid)
        if not c:
            continue
        done += 1
        row = c.get("row", {"代號": sid})
        # 未來估值:若快取 row 尚無(舊快取),即時用 eps_year 補算(免重抓財報)
        ey = c.get("eps_year")
        if "未來估值" not in row and ey:
            ev = [ey[y] for y in sorted(ey)]
            if len(ev) >= 2:
                e5 = ((ev[-1]/ev[0])**(1/(len(ev)-1))-1)*100 if ev[0] > 0 and ev[-1] > 0 else np.nan
                e3 = ((ev[-1]/ev[-3])**0.5-1)*100 if len(ev) >= 3 and ev[-3] > 0 and ev[-1] > 0 else np.nan
                cyc = (min(ev) <= 0) or any(ev[i] < ev[i-1]*0.8 for i in range(1, len(ev)))
                for k, vv in forward_metrics(row.get("收盤"), row.get("近四季EPS"),
                                             row.get("PER(自算)"), e3, e5,
                                             row.get("最新月營收年增%"), cyc).items():
                    row[k] = vv
        # 合理價鬧鐘:舊快取若無,從 hist 的 PE_Pxx × forward EPS 補算(免重抓財報)
        hist_c = c.get("hist", {"代號": sid})
        if "合理價" not in row:
            fwd_eps = row.get("預估明年EPS") or row.get("近四季EPS")
            if pd.notna(fwd_eps) and fwd_eps and fwd_eps > 0:
                for q, label in (("PE_P50", "合理價"), ("PE_P25", "偏便宜價"), ("PE_P10", "深度買點價")):
                    pe_q = hist_c.get(q)
                    if pe_q is not None and pd.notna(pe_q):
                        row[label] = round(float(fwd_eps) * float(pe_q), 1)
        rows.append(row)
        hists.append(hist_c)
        name = namemap.get(sid, sid)
        if c.get("rev_year"): rev_years[f"{sid} {name}"] = c["rev_year"]
        if ey: eps_years[f"{sid} {name}"] = ey
        if c.get("q_gm"):     q_gms[f"{sid} {name}"] = c["q_gm"]
    if not rows:
        print("尚無任何已完成資料,略過輸出"); return 0

    df = pd.DataFrame(rows)
    cols = ["代號", "名稱", "金融", "最新季",
            "5年營收CAGR%", "5年平均淨利率%", "5年平均ROE%",
            "毛利率%", "營益率%", "淨利率%", "近四季EPS", "近四季ROE%", "ROIC%(真實)",
            "負債比%", "流動比%", "存貨年增%", "獲利含金量", "近四季自由現金流(億)",
            "收盤", "市值(億)", "PER(自算)", "PE位階%", "PBR", "殖利率%", "最新月營收年增%",
            "成長率g%", "預估明年EPS", "ForwardPE", "ForwardPE保守", "PEG", "未來估值",
            "合理價", "偏便宜價", "深度買點價"]
    df = df[[c for c in cols if c in df.columns]]
    TEXT_COLS = ("代號", "名稱", "金融", "最新季", "未來估值")
    for col in df.columns:
        if col not in TEXT_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    sort_key = "5年平均ROE%" if "5年平均ROE%" in df.columns else "近四季ROE%"
    if sort_key in df.columns:
        df = df.sort_values(sort_key, ascending=False, na_position="last")

    def pivot_years(d):
        if not d:
            return pd.DataFrame()
        years = sorted({y for v in d.values() for y in v})
        out = pd.DataFrame({lbl: pd.Series(v) for lbl, v in d.items()}).T
        return out.reindex(columns=years)

    hdf = pd.DataFrame(hists)
    hcols = ["代號", "名稱", "金融"]
    for label in ("毛利率", "營益率", "淨利率", "ROE", "PER", "PBR", "殖利率"):
        hcols += [f"{label}現", f"{label}5年均", f"{label}位階%"]
    hdf = hdf[[c for c in hcols if c in hdf.columns]]
    if "代號" in df.columns and "代號" in hdf.columns:
        hdf = hdf.set_index("代號").reindex(df["代號"]).reset_index()

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="財報估值比較", index=False)
        hdf.to_excel(xw, sheet_name="相對歷史水位", index=False)
        ry, ey = pivot_years(rev_years), pivot_years(eps_years)
        if not ry.empty: ry.to_excel(xw, sheet_name="逐年營收(億)")
        if not ey.empty: ey.to_excel(xw, sheet_name="逐年EPS")
        if q_gms:                                       # 逐季毛利率(近8季,供拐點掃描算毛利拐頭)
            qg = pd.DataFrame({lbl: pd.Series(v) for lbl, v in q_gms.items()}).T
            qg = qg.reindex(columns=sorted(qg.columns))
            qg.to_excel(xw, sheet_name="逐季毛利率")
    print(f"  → 已更新 {OUTPUT}(目前 {done}/{len(PICKS)} 檔)")
    return done


# ---------- 主流程 ----------
def main():
    t0 = time.time()
    dl = make_loader()
    namemap = load_names(dl)                            # 代號→官方股名
    todo = [s for s in PICKS if load_cache(s) is None]
    print(f"總 {len(PICKS)} 檔,已完成 {len(PICKS)-len(todo)} 檔,待抓 {len(todo)} 檔")
    build_output(namemap)                              # 先用既有快取出一版(確保隨時有進度檔)

    for i, sid in enumerate(todo, 1):
        if time.time() - t0 > MAX_RUNTIME_MIN * 60:     # 整輪超時 → 存檔退出,下輪靠快取續抓
            print(f"⏲ 已達 {MAX_RUNTIME_MIN} 分上限,本輪先收尾(剩 {len(todo)-i+1} 檔下輪續抓)")
            break
        name = namemap.get(sid, sid)
        print(f"[{i}/{len(todo)}] 抓取 {sid} {name} ...")
        tries = 0
        while True:
            try:
                raw = fetch_one(dl, sid, START_DATE)
                row, perf, rev_y, eps_y, hist = summary_row(sid, name, raw)
                if sid in FINANCIALS:                  # 金融股標記
                    row["金融"] = "🏦"; hist["金融"] = "🏦"
                q3 = {}                                # 逐季毛利率(近8季,供拐點掃描算毛利拐頭)
                if not perf.empty and "毛利率%" in perf.columns:
                    qm = perf["毛利率%"].dropna().tail(8)
                    q3 = {str(d)[:10]: round(float(x), 2) for d, x in qm.items()}
                save_cache(sid, {"row": row, "hist": hist,
                                 "rev_year": rev_y, "eps_year": eps_y, "q_gm": q3})  # 立刻存,續跑用
                break
            except Exception as e:
                if _is_rate_limit(e) and tries < MAX_RATE_RETRY:
                    tries += 1
                    print(f"  ⏸ 疑似額度用罄 → 睡 {RATE_WAIT_CAP}s 短重試({tries}/{MAX_RATE_RETRY})")
                    time.sleep(RATE_WAIT_CAP); continue
                if _is_rate_limit(e):                  # 重試仍撞額度 → 跳過(不寫快取,下輪會再抓)
                    print(f"  ↷ {sid} 額度未恢復,本輪跳過(下輪靠快取補)")
                    break
                print(f"  ! {sid} 失敗:{e}")           # 真失敗(非額度)才寫空快取,避免一直重試壞票
                save_cache(sid, {"row": {"代號": sid, "名稱": name}, "hist": {"代號": sid, "名稱": name}})
                break
        if i % 10 == 0:                                # 每 10 檔重建一次 Excel(分段保存進度)
            build_output(namemap)
        time.sleep(RATE_SLEEP)

    n = build_output(namemap)                          # 收尾再出一版完整的
    print(f"\n完成 → {OUTPUT}({n}/{len(PICKS)} 檔)")


if __name__ == "__main__":
    main()
