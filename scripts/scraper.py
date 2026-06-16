"""
千張大戶週爬蟲 — GitHub Actions 版
資料來源：https://stock.wearn.com/holders.asp
"""

import os
import sys
import time
import datetime
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pathlib import Path

# ────────────────────────────────────────────
# 設定（也可從環境變數注入）
# ────────────────────────────────────────────

DEFAULT_STOCKS = [
    '2330',  # 台積電
    '2317',  # 鴻海
    '2454',  # 聯發科
    '2308',  # 台達電
    '2382',  # 廣達
    '3711',  # 日月光投控
    '2449',  # 京元電子
    '6669',  # 緯穎
    '3017',  # 奇鋐
    '2383',  # 台光電
]

# 從環境變數讀取（GitHub Actions 手動觸發時可覆蓋）
env_stocks = os.environ.get('INPUT_STOCKS', '')
STOCKS = [s.strip() for s in env_stocks.split(',')] if env_stocks else DEFAULT_STOCKS

WEEKS = int(os.environ.get('INPUT_WEEKS', '8'))
SLEEP = 1.5
OUTPUT_DIR = Path('data')
OUTPUT_DIR.mkdir(exist_ok=True)

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Referer': 'https://stock.wearn.com/',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}

TIER_LABELS = {
    '1000張以上': '千張大戶',
    '800~1000':  '800~1000張',
    '600~800':   '600~800張',
    '400~600':   '400~600張',
    '200~400':   '200~400張',
    '100~200':   '100~200張',
    '50~100':    '50~100張',
    '40~50':     '40~50張',
    '30~40':     '30~40張',
    '20~30':     '20~30張',
    '15~20':     '15~20張',
    '10~15':     '10~15張',
    '5~10':      '5~10張',
    '1~5':       '1~5張',
    '1張以下':    '零股散戶',
}


# ────────────────────────────────────────────
# 工具函數
# ────────────────────────────────────────────

def gen_fridays(n: int) -> list[str]:
    """產生最近 n 個週五的民國日期字串（從最新到最舊）"""
    today = datetime.date.today()
    days_back = (today.weekday() - 4) % 7
    last_fri = today - datetime.timedelta(days=days_back)
    result = []
    for i in range(n):
        d = last_fri - datetime.timedelta(weeks=i)
        roc = d.year - 1911
        result.append(f"{roc}{d.month:02d}{d.day:02d}")
    return result


def roc_to_date(roc_str: str) -> str:
    roc = int(roc_str[:3])
    m, d = roc_str[3:5], roc_str[5:7]
    return f"{roc + 1911}/{m}/{d}"


def fetch_holders(stock_id: str, date_str: str) -> list[dict]:
    url = f"https://stock.wearn.com/holders.asp?kind={stock_id}&d={date_str}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.encoding = 'big5'
    except Exception as e:
        print(f"    ❌ 連線失敗: {e}")
        return []

    if r.status_code != 200:
        print(f"    ⚠️  HTTP {r.status_code}")
        return []

    soup = BeautifulSoup(r.text, 'html.parser')

    # 公司名稱
    title = soup.title.string if soup.title else ''
    company = title.split('(')[0].strip() if '(' in title else stock_id

    table = soup.find('table')
    if not table:
        return []

    rows, results = table.find_all('tr'), []
    for row in rows:
        cols = [td.get_text(strip=True) for td in row.find_all('td')]
        if len(cols) < 4 or '人' not in cols[1]:
            continue
        try:
            chg_pct = float(cols[4]) if len(cols) > 4 else 0.0
            results.append({
                '日期':       roc_to_date(date_str),
                '代號':       stock_id,
                '名稱':       company,
                '分級':       TIER_LABELS.get(cols[2], cols[2]),
                '持股人數':    int(cols[1].replace('人','').replace(',','')),
                '週增減人數':  int(cols[0].replace(',','').replace('+','')),
                '持股比%':    float(cols[3].replace('%','')),
                '持股比增減%': chg_pct,
            })
        except (ValueError, IndexError):
            continue
    return results


# ────────────────────────────────────────────
# 主程式
# ────────────────────────────────────────────

def main():
    dates = gen_fridays(WEEKS)
    print(f"📋 股票清單: {STOCKS}")
    print(f"📅 日期範圍: {roc_to_date(dates[-1])} ～ {roc_to_date(dates[0])}（共{WEEKS}週）\n")

    all_rows, total = [], len(STOCKS) * len(dates)
    count = 0

    for stock_id in STOCKS:
        for date_str in dates:
            count += 1
            label = f"{stock_id} {roc_to_date(date_str)}"
            print(f"  [{count:>3}/{total}] {label} ...", end=' ', flush=True)
            rows = fetch_holders(stock_id, date_str)
            if rows:
                all_rows.extend(rows)
                # 只顯示千張大戶那行的重點數字
                big = next((r for r in rows if r['分級'] == '千張大戶'), None)
                if big:
                    chg = big['週增減人數']
                    sign = '+' if chg >= 0 else ''
                    print(f"✅ 千張大戶 {big['持股比%']:.2f}% ({sign}{chg}人)")
                else:
                    print(f"✅ {len(rows)} 筆")
            else:
                print("（無資料）")
            time.sleep(SLEEP)

    if not all_rows:
        print("\n❌ 沒有抓到任何資料")
        sys.exit(1)

    df = pd.DataFrame(all_rows)
    today_str = datetime.date.today().strftime('%Y%m%d')

    # ── 存 CSV（方便 git diff 追蹤變化）
    csv_path = OUTPUT_DIR / 'holders_all.csv'
    df.sort_values(['代號','日期','分級']).to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"\n💾 CSV 已存至 {csv_path}")

    # ── 千張大戶篩出
    big = df[df['分級'] == '千張大戶'].copy()
    name_map = df[['代號','名稱']].drop_duplicates().set_index('代號')['名稱']

    def add_name(pv):
        pv.insert(0, '名稱', pv.index.map(name_map))
        return pv

    pv_pct  = add_name(big.pivot_table(index='代號', columns='日期', values='持股比%',    aggfunc='first'))
    pv_cnt  = add_name(big.pivot_table(index='代號', columns='日期', values='持股人數',   aggfunc='first'))
    pv_chg  = add_name(big.pivot_table(index='代號', columns='日期', values='週增減人數', aggfunc='first'))
    pv_cpct = add_name(big.pivot_table(index='代號', columns='日期', values='持股比增減%', aggfunc='first'))

    # ── 存 Excel
    xlsx_path = OUTPUT_DIR / f'千張大戶_{today_str}.xlsx'
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as w:
        pv_pct.to_excel(w,  sheet_name='持股比%趨勢')
        pv_cnt.to_excel(w,  sheet_name='持股人數趨勢')
        pv_chg.to_excel(w,  sheet_name='週增減人數')
        pv_cpct.to_excel(w, sheet_name='持股比增減%')
        df.sort_values(['代號','日期','分級']).to_excel(w, sheet_name='完整原始資料', index=False)

    print(f"📊 Excel 已存至 {xlsx_path}")

    # ── 終端機摘要
    print("\n" + "="*60)
    print(f"{'代號':<6} {'名稱':<10} {'最新持股比%':>10} {'週增減人數':>10} {'4週趨勢':>10}")
    print("-"*60)
    latest_date = dates[0]
    latest_label = roc_to_date(latest_date)
    for stock_id in STOCKS:
        sub = big[big['代號'] == stock_id].sort_values('日期')
        if sub.empty:
            continue
        latest = sub[sub['日期'] == latest_label]
        if latest.empty:
            continue
        r = latest.iloc[0]
        # 4週趨勢箭頭
        pcts = sub['持股比%'].tail(4).tolist()
        trend = ''.join(['↑' if pcts[i] > pcts[i-1] else '↓'
                         for i in range(1, len(pcts))])
        chg = r['週增減人數']
        sign = '+' if chg >= 0 else ''
        print(f"{stock_id:<6} {r['名稱']:<10} {r['持股比%']:>9.2f}% "
              f"{sign}{chg:>8}人  {trend:>10}")
    print("="*60)
    print(f"\n✅ 完成！共 {len(df)} 筆資料")


if __name__ == '__main__':
    main()
