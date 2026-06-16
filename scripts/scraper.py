"""
千張大戶週爬蟲 — GitHub Actions 版 v2
修正：Big5編碼解析 + Session Cookie + 防爬 Header + 重試機制
"""

import os, sys, time, datetime, requests, pandas as pd
from bs4 import BeautifulSoup
from pathlib import Path

# ── 設定 ──────────────────────────────────────────────────────
DEFAULT_STOCKS = [
    '2330','2317','2454','2308','2382',
    '3711','2449','6669','3017','2383',
]
env_stocks    = os.environ.get('INPUT_STOCKS', '')
STOCKS        = [s.strip() for s in env_stocks.split(',')] if env_stocks else DEFAULT_STOCKS
WEEKS         = int(os.environ.get('INPUT_WEEKS', '8'))
SLEEP         = 2.0
OUTPUT_DIR    = Path('data')
OUTPUT_DIR.mkdir(exist_ok=True)

TIER_LABELS = {
    '1000張以上':'千張大戶','800~1000':'800~1000張','600~800':'600~800張',
    '400~600':'400~600張','200~400':'200~400張','100~200':'100~200張',
    '50~100':'50~100張','40~50':'40~50張','30~40':'30~40張',
    '20~30':'20~30張','15~20':'15~20張','10~15':'10~15張',
    '5~10':'5~10張','1~5':'1~5張','1張以下':'零股散戶',
}

# ── Session（帶 Cookie）───────────────────────────────────────
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/125.0.0.0 Safari/537.36'
        ),
        'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection':      'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control':   'max-age=0',
    })
    # 先打首頁取 Cookie
    try:
        s.get('https://stock.wearn.com/', timeout=10)
        time.sleep(1)
    except Exception:
        pass
    return s

SESSION = make_session()

# ── 工具 ─────────────────────────────────────────────────────
def gen_fridays(n: int) -> list[str]:
    today       = datetime.date.today()
    days_back   = (today.weekday() - 4) % 7
    last_fri    = today - datetime.timedelta(days=days_back)
    result = []
    for i in range(n):
        d = last_fri - datetime.timedelta(weeks=i)
        result.append(f"{d.year - 1911}{d.month:02d}{d.day:02d}")
    return result

def roc_to_date(s: str) -> str:
    return f"{int(s[:3])+1911}/{s[3:5]}/{s[5:7]}"

# ── 核心爬取 ──────────────────────────────────────────────────
def fetch_holders(stock_id: str, date_str: str, retry: int = 3) -> list[dict]:
    url = f"https://stock.wearn.com/holders.asp?kind={stock_id}&d={date_str}"

    for attempt in range(1, retry + 1):
        try:
            r = SESSION.get(url, timeout=15)
        except Exception as e:
            print(f"    ⚠️  連線失敗({attempt}/{retry}): {e}")
            time.sleep(3)
            continue

        if r.status_code != 200:
            print(f"    ⚠️  HTTP {r.status_code} ({attempt}/{retry})")
            time.sleep(3)
            continue

        # ── Big5 解碼（關鍵修正）
        # 優先用 chardet 偵測，fallback 用 big5
        try:
            import chardet
            detected = chardet.detect(r.content)
            enc = detected.get('encoding') or 'big5'
        except ImportError:
            enc = 'big5'

        try:
            html = r.content.decode(enc, errors='replace')
        except Exception:
            html = r.content.decode('big5', errors='replace')

        soup = BeautifulSoup(html, 'html.parser')

        # 公司名稱
        title   = soup.title.string if soup.title else ''
        company = title.split('(')[0].strip() if '(' in title else stock_id

        # 找表格（聚財網的資料表是第一個含 td 的 table）
        tables = soup.find_all('table')
        data_table = None
        for t in tables:
            if t.find('td'):
                data_table = t
                break

        if not data_table:
            print(f"    ⚠️  找不到表格（可能被擋或頁面結構變更）")
            # debug：印出前500字看看
            print(f"    DEBUG html[:300]: {html[:300]}")
            break

        rows, results = data_table.find_all('tr'), []
        for row in rows:
            cols = [td.get_text(strip=True) for td in row.find_all('td')]
            if len(cols) < 4:
                continue
            # 判斷是否為資料列：第2欄含「人」字
            if '人' not in cols[1]:
                continue
            try:
                chg_pct = float(cols[4]) if len(cols) > 4 else 0.0
                results.append({
                    '日期':       roc_to_date(date_str),
                    '代號':       stock_id,
                    '名稱':       company,
                    '分級':       TIER_LABELS.get(cols[2], cols[2]),
                    '持股人數':    int(cols[1].replace('人','').replace(',','')),
                    '週增減人數':  int(cols[0].replace(',','').replace('+','').replace('\xa0','')),
                    '持股比%':    float(cols[3].replace('%','')),
                    '持股比增減%': chg_pct,
                })
            except (ValueError, IndexError):
                continue

        return results  # 成功就回傳（即使是空 list）

    return []  # 全部 retry 失敗

# ── 主程式 ────────────────────────────────────────────────────
def main():
    dates = gen_fridays(WEEKS)
    print(f"📋 股票清單: {STOCKS}")
    print(f"📅 日期範圍: {roc_to_date(dates[-1])} ～ {roc_to_date(dates[0])}（共{WEEKS}週）\n")

    all_rows, total, count = [], len(STOCKS) * len(dates), 0

    for stock_id in STOCKS:
        for date_str in dates:
            count += 1
            print(f"  [{count:>3}/{total}] {stock_id} {roc_to_date(date_str)} ...", end=' ', flush=True)
            rows = fetch_holders(stock_id, date_str)
            if rows:
                all_rows.extend(rows)
                big = next((r for r in rows if r['分級'] == '千張大戶'), None)
                if big:
                    chg  = big['週增減人數']
                    sign = '+' if chg >= 0 else ''
                    print(f"✅  千張大戶 {big['持股比%']:.2f}% ({sign}{chg}人)")
                else:
                    print(f"✅  {len(rows)} 筆（無千張大戶欄）")
            else:
                print("（無資料）")
            time.sleep(SLEEP)

    if not all_rows:
        print("\n❌ 沒有抓到任何資料")
        sys.exit(1)

    df        = pd.DataFrame(all_rows)
    today_str = datetime.date.today().strftime('%Y%m%d')
    big       = df[df['分級'] == '千張大戶'].copy()
    name_map  = df[['代號','名稱']].drop_duplicates().set_index('代號')['名稱']

    def pv(col):
        p = big.pivot_table(index='代號', columns='日期', values=col, aggfunc='first')
        p.insert(0, '名稱', p.index.map(name_map))
        return p

    # CSV（累積，git diff 看變化）
    csv_path = OUTPUT_DIR / 'holders_all.csv'
    if csv_path.exists():
        old = pd.read_csv(csv_path, dtype=str)
        df  = pd.concat([old, df.astype(str)]).drop_duplicates(
                  subset=['日期','代號','分級'], keep='last')
    df.sort_values(['代號','日期','分級']).to_csv(csv_path, index=False, encoding='utf-8-sig')

    # Excel
    xlsx_path = OUTPUT_DIR / f'千張大戶_{today_str}.xlsx'
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as w:
        pv('持股比%'   ).to_excel(w, sheet_name='持股比%趨勢')
        pv('持股人數'   ).to_excel(w, sheet_name='持股人數趨勢')
        pv('週增減人數' ).to_excel(w, sheet_name='週增減人數')
        pv('持股比增減%').to_excel(w, sheet_name='持股比增減%')
        df.to_excel(w, sheet_name='完整原始資料', index=False)

    print(f"\n💾 CSV  → {csv_path}")
    print(f"📊 Excel→ {xlsx_path}")

    # 終端機摘要
    print("\n" + "="*65)
    print(f"{'代號':<6} {'名稱':<10} {'最新持股比%':>10} {'週增減':>8} {'近4週趨勢':>10}")
    print("-"*65)
    latest = roc_to_date(dates[0])
    for sid in STOCKS:
        sub = big[big['代號']==sid].sort_values('日期')
        if sub.empty: continue
        row = sub[sub['日期']==latest]
        if row.empty: continue
        r = row.iloc[0]
        pcts  = sub['持股比%'].tail(4).tolist()
        trend = ''.join(['↑' if pcts[i]>pcts[i-1] else '↓' for i in range(1,len(pcts))])
        chg   = r['週增減人數']
        sign  = '+' if int(chg)>=0 else ''
        print(f"{sid:<6} {r['名稱']:<10} {float(r['持股比%']):>9.2f}%  "
              f"{sign}{int(chg):>6}人  {trend:>10}")
    print("="*65)
    print(f"\n✅ 完成！共 {len(df)} 筆資料")

if __name__ == '__main__':
    main()
