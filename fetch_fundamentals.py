#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股「南亞科等級」基本資料抓取器
================================
抓取季度損益/EPS、月營收、毛利率等基本面數據，輸出成 Excel。

兩條路線：
  A) FinMind  -- 推薦。官方資料的免費 API 封裝，乾淨、合法、不會被封 IP。
  B) Goodinfo -- 你原本想抓的站。可行但脆弱(反爬、需瀏覽器標頭、版型常改)，
                 且其服務條款不鼓勵自動抓取，請務必降速、僅供個人研究。

預設標的 = 我建議優先驗證的同循環族群：
  2408 南亞科, 2344 華邦電, 6173 信昌電, 6770 力積電, 3090 日電貿, 3231 緯創

執行前安裝套件：
  pip install finmind pandas openpyxl requests lxml
"""

import time
import pandas as pd

TICKERS = ["3481", "2327", "6669", "2379", "1560", "8210"]
START_DATE = "2024-01-01"          # 抓兩年, 足以看出循環轉折
OUTPUT = "台股基本面.xlsx"


# ============================================================
# 路線 A：FinMind（推薦）
# 註冊免費 token：https://finmindtrade.com/  (免 token 也能跑，但有額度限制)
# ============================================================
def fetch_finmind(tickers, start_date, token=""):
    from FinMind.data import DataLoader
    dl = DataLoader()
    if token:
        dl.login_by_token(api_token=token)

    out = {}  # ticker -> {表名: DataFrame}
    for sid in tickers:
        print(f"[FinMind] 抓取 {sid} ...")
        try:
            income = dl.taiwan_stock_financial_statement(stock_id=sid, start_date=start_date)
            revenue = dl.taiwan_stock_month_revenue(stock_id=sid, start_date=start_date)
            price = dl.taiwan_stock_daily(stock_id=sid, start_date=start_date)
            out[sid] = {"季度損益": income, "月營收": revenue, "日價量": price}
        except Exception as e:
            print(f"  ! {sid} 失敗: {e}")
        time.sleep(1.0)   # 對 API 客氣一點
    return out


def derive_metrics_finmind(income: pd.DataFrame) -> pd.DataFrame:
    """從 FinMind 損益表長表，整理出我們在南亞科分析用到的核心指標。"""
    if income is None or income.empty:
        return pd.DataFrame()
    piv = income.pivot_table(index="date", columns="type", values="value", aggfunc="first")
    want = {
        "Revenue": "營業收入",
        "GrossProfit": "營業毛利",
        "OperatingIncome": "營業利益",
        "IncomeAfterTaxes": "稅後淨利",
        "EPS": "EPS",
    }
    have = {k: v for k, v in want.items() if k in piv.columns}
    m = piv[list(have)].rename(columns=have).copy()
    if "營業收入" in m and "營業毛利" in m:
        m["毛利率%"] = (m["營業毛利"] / m["營業收入"] * 100).round(2)
    if "營業收入" in m and "營業利益" in m:
        m["營業利益率%"] = (m["營業利益"] / m["營業收入"] * 100).round(2)
    return m.sort_index()


# ============================================================
# 路線 B：Goodinfo（你原本的網址；脆弱，請降速）
# ============================================================
def fetch_goodinfo(stock_id, rpt_cat="XX_M_QUAR_ACC"):
    """
    rpt_cat 常用值：
      XX_M_QUAR_ACC  季度累計財報   XX_M_QUAR  單季財報
      XX_M_YEAR      年度財報
    Goodinfo 會把表格 server-side 放進 HTML，關鍵是要帶『瀏覽器標頭』，
    否則(像 web 抓取工具那樣)會被回空白頁。
    """
    import requests
    url = "https://goodinfo.tw/tw/StockFinDetail.asp"
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0 Safari/537.36"),
        "Referer": "https://goodinfo.tw/tw/StockList.asp",
        "Accept-Language": "zh-TW,zh;q=0.9",
    }
    params = {"RPT_CAT": rpt_cat, "STOCK_ID": stock_id}
    r = requests.get(url, params=params, headers=headers, timeout=20)
    r.encoding = "utf-8"               # 新版 goodinfo 是 utf-8（舊版曾是 big5）
    # 財報表格 id 多為 'tblFinDetail'；先試精準抓，失敗再退而求其次抓全部表格
    try:
        tables = pd.read_html(r.text, attrs={"id": "tblFinDetail"})
        return tables[0]
    except Exception:
        tables = pd.read_html(r.text)          # 抓頁面所有表格
        # 通常最大的那張就是財報主表
        return max(tables, key=lambda t: t.shape[0] * t.shape[1]) if tables else pd.DataFrame()


def run_goodinfo(tickers):
    out = {}
    for sid in tickers:
        print(f"[Goodinfo] 抓取 {sid} ...")
        try:
            out[sid] = fetch_goodinfo(sid)
        except Exception as e:
            print(f"  ! {sid} 失敗: {e}")
        time.sleep(8.0)    # 重要：goodinfo 會封過快的 IP，務必慢
    return out


# ============================================================
# 輸出
# ============================================================
def export(data_by_ticker, path, derive=None):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for sid, payload in data_by_ticker.items():
            if isinstance(payload, dict):           # FinMind 多表
                for name, df in payload.items():
                    if df is not None and not df.empty:
                        df.to_excel(xw, sheet_name=f"{sid}_{name}"[:31], index=False)
                if derive and "季度損益" in payload:
                    m = derive(payload["季度損益"])
                    if not m.empty:
                        m.to_excel(xw, sheet_name=f"{sid}_核心指標"[:31])
            else:                                   # Goodinfo 單表
                if payload is not None and not payload.empty:
                    payload.to_excel(xw, sheet_name=str(sid)[:31])
    print(f"\n已輸出：{path}")


if __name__ == "__main__":
    # ---- 預設走 FinMind（推薦）。若沒裝 FinMind 會自動退回 Goodinfo。----
    try:
        data = fetch_finmind(TICKERS, START_DATE, token="")   # 有 token 填這裡
        export(data, OUTPUT, derive=derive_metrics_finmind)
    except ImportError:
        print("未安裝 FinMind，改用 Goodinfo（請確保已 pip install lxml）...")
        data = run_goodinfo(TICKERS)
        export(data, OUTPUT)
