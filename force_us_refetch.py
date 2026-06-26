# -*- coding: utf-8 -*-
"""
強制重抓美股 A 級體檢 force_us_refetch.py
=======================================================================
patch_us_profile 在 CI 一直死鎖(profile 端點某些股 hang) → 改用穩定路徑:
從體檢總表刪掉 A 級 174 列(其他 7700 維持原狀), 然後觸發 us-fundamentals.yml
這時這 174 檔自動進「待抓」清單, 用已修好的 marketCap/dividendYield 碼重抓。

僅 174 檔 × 5 endpoint = 870 calls, FMP 付費 ~3 分內完成。
"""
import os
import pandas as pd

SRC = "data/美股體檢總表.xlsx"


def main():
    xls = pd.ExcelFile(SRC)
    sheets = {sh: pd.read_excel(SRC, sheet_name=sh) for sh in xls.sheet_names}
    h = sheets["體檢總表"]
    h["代號"] = h["代號"].astype(str)
    a_codes = h[h["評等"] == "A"]["代號"].tolist()
    print(f"刪除 A 級 {len(a_codes)} 列,讓 us-fundamentals 重抓")
    sheets["體檢總表"] = h[h["評等"] != "A"].copy()
    print(f"剩餘 {len(sheets['體檢總表'])} 列(原 {len(h)})")
    # 寫回
    tmp = SRC + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for sh, df in sheets.items():
            df.to_excel(xw, sheet_name=sh, index=False)
    os.replace(tmp, SRC)
    print(f"→ 已更新 {SRC}(A級已移除,等待 us-fundamentals 重抓填回)")


if __name__ == "__main__":
    main()
