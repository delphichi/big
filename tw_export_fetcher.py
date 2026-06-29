# -*- coding: utf-8 -*-
"""
台灣資通訊出口抓取器 tw_export_fetcher.py
=======================================================================
台灣月出口 = 全球科技股 3 週領先指標
財政部每月 7 號上午 10:00 公布前一月數據

三層 fallback:
  1. 財政部統計處 OpenData(主源,完整細分)
  2. OECD SDMX-JSON(備援,只有總值)
  3. 寫死 known historical(最後一道)

跑法獨立測試:
  python tw_export_fetcher.py
"""
import os
import re
import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone

TPE = timezone(timedelta(hours=8))


def from_mof_opendata():
    """財政部統計處 OpenData — 用 'sym' 編號抓 XML
    sym 編號參考:https://web02.mof.gov.tw/njswww/WebMain.aspx
      4220000000000022070 = 出口貿易值(月)
      4220000000000022300 = 進口貿易值(月)
    """
    # 抓出口貿易值總表(月)
    try:
        url = "https://web02.mof.gov.tw/njswww/webproxy.aspx"
        r = requests.get(url, params={
            "sym": "4220000000000022070",  # 出口總值(月,百萬美元)
            "run": "Y", "open": "1", "funid": "defjsptg9b"
        }, timeout=20)
        if r.status_code != 200: return None
        # XML 解析(財政部用 .NET 預設格式)
        text = r.text
        # 抓出最新一筆月份 + 出口總值
        # 格式類似 <Table><月份>113年5月</月份><出口總值>4xxxx</出口總值>...
        rows = re.findall(r"<Table>(.*?)</Table>", text, re.DOTALL)
        if not rows: return None
        latest = rows[-1] if rows else None
        if not latest: return None
        month = re.search(r"<月份>(.*?)</月份>", latest)
        export = re.search(r"<出口總值>([\d.,]+)</出口總值>", latest)
        if not (month and export): return None
        return {
            "源": "財政部 OpenData",
            "月份": month.group(1),
            "出口總值(百萬美元)": float(export.group(1).replace(",", "")),
        }
    except Exception as e:
        return None


def from_oecd():
    """OECD SDMX-JSON — 國際標準,穩定
    Dataset: MEI_TRD (Main Economic Indicators - Trade)
    台灣代碼: TWN, 出口指標: XTEXVA01
    """
    try:
        url = "https://stats.oecd.org/SDMX-JSON/data/MEI_TRD/TWN.XTEXVA01.IXOBSA.M/all"
        r = requests.get(url, params={"dimensionAtObservation": "allDimensions"},
                         timeout=20, headers={"Accept": "application/json"})
        if r.status_code != 200: return None
        j = r.json()
        # 解析 SDMX-JSON
        obs = j.get("dataSets", [{}])[0].get("observations", {})
        time_dim = j.get("structure", {}).get("dimensions", {}).get("observation", [])
        # 找 TIME_PERIOD index
        time_idx = next((i for i, d in enumerate(time_dim) if d.get("id")=="TIME_PERIOD"), -1)
        if time_idx < 0 or not obs: return None
        times = time_dim[time_idx].get("values", [])
        # 取最新一筆
        latest_key = max(obs.keys(), key=lambda k: int(k.split(":")[time_idx]))
        latest_val = obs[latest_key][0]
        latest_time_idx = int(latest_key.split(":")[time_idx])
        latest_time = times[latest_time_idx].get("id", "")
        return {
            "源": "OECD SDMX",
            "月份": latest_time,
            "出口指數(2015=100)": round(latest_val, 1),
        }
    except Exception:
        return None


def from_mof_news():
    """爬財政部新聞稿(備援第二道)
    財政部每月 7-8 號公布:
      https://www.mof.gov.tw/multiplehtml/list/4 (新聞清單)
    抓最新一篇「進出口貿易」相關新聞,parse 出總值"""
    try:
        url = "https://www.mof.gov.tw/list/4"
        r = requests.get(url, timeout=20)
        if r.status_code != 200: return None
        # 找含「進出口」的新聞連結(粗略 regex)
        m = re.search(r'href="(/multiplehtml/[^"]+)"[^>]*>[^<]*?進出口[^<]*?\d+月', r.text)
        if not m: return None
        return {"源": "財政部新聞稿", "原文連結": f"https://www.mof.gov.tw{m.group(1)}",
                "註": "需點進去看詳細數字"}
    except Exception:
        return None


def yoy_calc(values):
    """從一串月度數據算 YoY%"""
    if not values or len(values) < 13: return None
    cur = values[-1]
    yr = values[-13]
    if not yr or yr == 0: return None
    return round((cur/yr - 1) * 100, 2)


def fetch_all():
    """完整流程:三源試,優先 MOF → OECD → news"""
    today = datetime.now(TPE).strftime("%Y-%m-%d")
    print(f"\n=== 台灣資通訊出口抓取 {today} ===\n")

    results = []
    for fn, name in [(from_mof_opendata, "財政部 OpenData"),
                     (from_oecd, "OECD SDMX"),
                     (from_mof_news, "財政部新聞稿")]:
        print(f"嘗試 {name} ...", end=" ")
        d = fn()
        if d:
            print("✅"); results.append(d)
            for k, v in d.items():
                print(f"  {k}: {v}")
        else:
            print("❌")

    if not results:
        print("\n⚠️ 三源都失敗,建議手動查:")
        print("  https://web02.mof.gov.tw/njswww/WebMain.aspx")
        print("  搜尋『進出口貿易初步統計』")
        return None
    return results


if __name__ == "__main__":
    fetch_all()
