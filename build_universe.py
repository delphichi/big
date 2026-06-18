#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
能力圈股池建立器  build_universe.py
==================================
從 FinMind TaiwanStockInfo 抓全台股清單,依「產業別」篩出能力圈標的,寫成 universe.txt,
供 total_screener.py 讀取(逐檔深挖五維)。

  能力圈 → FinMind 產業別關鍵字(用子字串比對,容許上市/上櫃命名差異):
    半導體        → 半導體
    網通          → 通信網路
    被動元件/零組件 → 電子零組件
    PC/伺服器代工  → 電腦及週邊
    金融銀行證券   → 金融 / 證券 / 銀行
  排除:非 4 碼普通股、00 開頭(ETF)、名稱含 ETF/購/售/期(權證/期貨)。

用法(需 FINMIND_TOKEN,見 README):
  python3 build_universe.py                       # 預設全部能力圈關鍵字
  python3 build_universe.py --keywords 半導體 金融  # 只要某幾類
  python3 build_universe.py --out universe.txt
"""
import os, re, argparse, datetime as _dt

# 能力圈 → 產業別關鍵字(子字串比對)
CIRCLE_KEYWORDS = ["半導體", "通信網路", "電子零組件", "電腦及週邊", "金融", "證券", "銀行"]

_CODE = re.compile(r"^[1-9]\d{3}$")        # 1xxx–9xxx 普通股;排除 00 開頭 ETF
_BADNAME = re.compile(r"(ETF|ETN|購\d|售\d|期貨|存託|受益)")


def filter_universe(records, keywords):
    """records: 可迭代的 (stock_id, stock_name, industry_category)。回傳篩後 list[dict],已去重排序。
    純函式、不依賴 pandas,方便離線測試。"""
    seen, out = set(), []
    for sid, name, ind in records:
        sid = str(sid).strip()
        name = str(name or "").strip()
        ind = str(ind or "").strip()
        if not _CODE.match(sid) or sid in seen:
            continue
        if _BADNAME.search(name):
            continue
        if not any(k in ind for k in keywords):
            continue
        seen.add(sid)
        out.append({"id": sid, "name": name, "industry": ind})
    out.sort(key=lambda r: r["id"])
    return out


def fetch_records():
    """從 FinMind 抓 TaiwanStockInfo,轉成 (id, name, industry) 串列。"""
    from FinMind.data import DataLoader
    dl = DataLoader()
    token = os.environ.get("FINMIND_TOKEN", "")
    if token:
        dl.login_by_token(api_token=token)
    df = dl.taiwan_stock_info()
    cols = df.columns
    name_col = "stock_name" if "stock_name" in cols else "name"
    ind_col = "industry_category" if "industry_category" in cols else "industry"
    return list(zip(df["stock_id"], df[name_col], df[ind_col]))


def write_universe(rows, path, keywords):
    by_ind = {}
    for r in rows:
        by_ind[r["industry"]] = by_ind.get(r["industry"], 0) + 1
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"# 能力圈股池  共 {len(rows)} 檔  產生於 {_dt.date.today().isoformat()}\n")
        f.write(f"# 關鍵字:{'、'.join(keywords)}\n")
        f.write("# 各產業檔數:" + "、".join(f"{k}={v}" for k, v in sorted(by_ind.items())) + "\n")
        f.write("# 格式:代號  # 名稱｜產業(total_screener 只讀第一欄代號)\n")
        for r in rows:
            f.write(f"{r['id']}  # {r['name']}｜{r['industry']}\n")
    return by_ind


def main():
    ap = argparse.ArgumentParser(description="能力圈股池建立器")
    ap.add_argument("--keywords", nargs="*", default=CIRCLE_KEYWORDS, help="產業別關鍵字(子字串比對)")
    ap.add_argument("--out", default="universe.txt")
    args = ap.parse_args()

    try:
        records = fetch_records()
    except Exception as e:
        print(f"抓取 TaiwanStockInfo 失敗(需安裝 FinMind 並設 FINMIND_TOKEN):{e}")
        return
    rows = filter_universe(records, args.keywords)
    by_ind = write_universe(rows, args.out, args.keywords)
    print(f"已寫入 {args.out}:共 {len(rows)} 檔")
    for k, v in sorted(by_ind.items()):
        print(f"  {k}: {v}")
    print(f"\n接著:python3 total_screener.py  (會自動讀 {args.out})")


if __name__ == "__main__":
    main()
