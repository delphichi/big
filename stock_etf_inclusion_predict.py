# -*- coding: utf-8 -*-
"""
0050 + 富櫃50 納入預測引擎 — 在 ETF 強迫買入前先卡位
=======================================================================
邏輯:
  0050 (臺灣50指數) 真實納入規則(FTSE Russell,季審 3/6/9/12 月第三個週五):
    ‧ 新納入:非成分股市值排名進入 **前 40 名** → 取得納入資格
    ‧ 候補名單:排名 **51~60 名** → 依市值遞補(有成分股掉出才補進)
    ‧ 剔除緩衝:現有成分股跌出 **前 61 名** 才剔除(避免頻繁換股)
    ‧ 快速納入:IPO/轉上市股市值排名 **前 20 名** → 5 日內直接納入,替換最後一名
  被納入 = 被動資金強迫買;等公告價格已反映。本腳本在「rank 41-60 卡位區 + 拐點
  + 估值未爆」時先標記,搶在 ETF 公告前進場;另標「成分股逼近 61 名」= 被動賣壓。
  另加「OTC 轉上市觀察」:上櫃大型股若轉上市(市值前20=快速納入)即可直接卡進。

資料 :
  data/twse_marketcap_weight.csv — TWSE 全市場排名 + 市值佔大盤比重%
                                    (上市股票口徑,正是 0050 的選股池;
                                     上櫃股如 5274 信驊 已自動排除)
  data/台股財報估值.xlsx [財報估值比較] — 我們 universe 的市值/PER/PBR/體質
  data/台股_體檢總表.xlsx [體檢總表]     — 評等/品質總分/估值/含金量/循環
  data/台股_拐點掃描.xlsx [全部訊號]     — 改善訊號數/分級(optional)

關鍵門檻(2026/06 snapshot):
  rank 50 = 6770 力積電 0.2577% ← 0050 邊緣
  rank 90 = 1605 華新   0.1181% ← 候選下限

評分流程:
  1. 把全市場 rank 100 對齊到我們 universe(用代號 join)
  2. 候選 A = 我們已體檢 + rank 41-60(真實卡位區:前40納入線~候補尾)
  3. 候選 B = blind spot — rank 41-60 但不在我們 PICKS → 該加進去抓
  4. 加分:評等 A/B + 估值便宜或合理 + 未來估值(PEG<1/forward便宜) + 拐點 + 含金量
     ★ 未來估值加分:當下 PE 可能過熱,但用明年 EPS 算便宜(PEG<1)就加分救回 —
       這是「成長還沒被 price in」的卡位機會。與 PE位階(過去尺)互補。
"""
import os
import pandas as pd
import numpy as np

VAL = "data/台股財報估值.xlsx"
HEA = "data/台股_體檢總表.xlsx"
INF = "data/台股_拐點掃描.xlsx"
WEIGHT = "data/twse_marketcap_weight.csv"
OTC = "data/otc_marketcap.csv"
OUT = "data/台股_0050納入預測.xlsx"

# 已知 0050 成分股(2025 名單)— 用於 universe 識別
KNOWN_0050 = {
    "2330", "2454", "2317", "2308", "3711", "2382", "2412", "3045", "2912",
    "3008", "6505", "2207", "2379", "2395", "2301", "2002", "1303", "1301",
    "2891", "2882", "2881", "2884", "2885", "2880", "2883", "2887", "2890",
    "2886", "2892", "5880", "2801", "5871",
    # 2026Q2 新納入
    "8046", "3443", "3665", "4958",
}


def main():
    if not os.path.exists(WEIGHT):
        print(f"找不到 {WEIGHT},無法做排名預測"); return
    rank = pd.read_csv(WEIGHT, dtype={"代號": str})
    THRESH_40 = float(rank[rank["rank"] == 40]["比重%"].iloc[0])   # 新納入線
    THRESH_60 = float(rank[rank["rank"] == 60]["比重%"].iloc[0])   # 候補尾
    THRESH_61 = float(rank[rank["rank"] == 61]["比重%"].iloc[0])   # 剔除線
    print(f"TWSE 真實門檻:rank 40(納入) = {THRESH_40}% / 60(候補尾) = {THRESH_60}% / 61(剔除) = {THRESH_61}%")
    top60 = set(rank[rank["rank"] <= 60]["代號"])
    rank_lookup = dict(zip(rank["代號"], rank["rank"]))
    weight_lookup = dict(zip(rank["代號"], rank["比重%"]))
    name_lookup = dict(zip(rank["代號"], rank["名稱"]))

    if not os.path.exists(VAL):
        print(f"⚠️ 找不到 {VAL},只能輸出 blind spot,無法做評分")
        val = pd.DataFrame(columns=["代號", "名稱"])
    else:
        val = pd.read_excel(VAL, "財報估值比較")
        val["代號"] = val["代號"].astype(str)
        for c in ["市值(億)", "收盤", "PER(自算)", "PE位階%", "PBR", "殖利率%", "最新月營收年增%"]:
            if c in val.columns:
                val[c] = pd.to_numeric(val[c], errors="coerce")

    # ---- 合併體檢/拐點 ----
    # 注意:財報估值表(val)現已含 forward 欄(成長率g%/ForwardPE/PEG/未來估值),
    # 故體檢只取『評分專屬』欄,避免與 val 的 forward 同名 merge 衝突(_x/_y)。
    hea = pd.DataFrame()
    if os.path.exists(HEA):
        hea = pd.read_excel(HEA, "體檢總表")
        hea["代號"] = hea["代號"].astype(str)
        hea_cols = ["代號", "評等", "品質總分", "估值", "循環股", "含金量",
                    "⑪動態惡化扣分", "負債比%", "流動比%", "存貨年增%"]
        hea = hea[[c for c in hea_cols if c in hea.columns]]
    inf = pd.DataFrame()
    if os.path.exists(INF):
        try:
            inf = pd.read_excel(INF, "全部訊號")
            inf["代號"] = inf["代號"].astype(str)
            inf = inf[["代號", "改善訊號數", "分級"]]
        except Exception:
            pass

    base = val.copy()
    if "名稱" not in base.columns:
        base["名稱"] = base["代號"].map(name_lookup)
    base["TWSE排名"] = base["代號"].map(rank_lookup)
    base["大盤比重%"] = base["代號"].map(weight_lookup)
    if not hea.empty:
        base = base.merge(hea, on="代號", how="left")
    if not inf.empty:
        base = base.merge(inf, on="代號", how="left")
    base["是否0050"] = base["代號"].apply(lambda c: "✓" if c in KNOWN_0050 else "")

    # ---- OTC 富櫃50 (006201 元大富櫃50) 排名 ----
    otc_rank = pd.DataFrame()
    if os.path.exists(OTC):
        otc_rank = pd.read_csv(OTC, dtype={"代號": str})
        otc_lookup = dict(zip(otc_rank["代號"], otc_rank["rank"]))
        otc_name_lookup = dict(zip(otc_rank["代號"], otc_rank["名稱"]))
        otc_mcap_lookup = dict(zip(otc_rank["代號"], otc_rank["市值億"]))
        base["OTC排名"] = base["代號"].map(otc_lookup)
        base["OTC市值億"] = base["代號"].map(otc_mcap_lookup)
        OTC_R50 = float(otc_rank[otc_rank["rank"] == 50]["市值億"].iloc[0])
        OTC_R90 = float(otc_rank[otc_rank["rank"] == 90]["市值億"].iloc[0])
        print(f"OTC 富櫃50 門檻:rank 50 = {OTC_R50}億 / rank 90 = {OTC_R90}億")
    else:
        OTC_R50 = OTC_R90 = None

    # ---- 候選 A:我們 universe 內 + rank 41-60 卡位區 + 不在 0050 ----
    cand_a = base[
        base["TWSE排名"].notna() &
        (base["TWSE排名"] >= 41) &
        (base["TWSE排名"] <= 60) &
        (base["是否0050"] == "")
    ].copy()

    def forward_bonus(r):
        """未來估值加分:股票買的是未來。當下PE可能過熱,但若用明年EPS算便宜(PEG<1)
        則加分救回 — 這是『成長還沒被 price in』的卡位機會。與 PE位階(過去尺)互補。"""
        s = 0
        peg = pd.to_numeric(r.get("PEG"), errors="coerce")
        fv = str(r.get("未來估值", ""))
        if pd.notna(peg) and 0 < peg < 1:
            s += 20                                   # PEG<1:成長未反映,最強 forward 訊號
        elif "未來便宜" in fv:
            s += 15
        elif "未來合理" in fv:
            s += 8
        return s

    def score(r):
        s = 0
        g = str(r.get("評等", ""))
        if g == "A": s += 30
        elif g == "B": s += 20
        elif g == "C": s += 10
        v = str(r.get("估值", ""))
        if "便宜" in v: s += 25
        elif "合理" in v: s += 18
        elif "偏貴" in v: s += 5
        s += forward_bonus(r)                          # 未來估值(forward PE/PEG)加分
        # ⑪ 動態惡化扣分:乘 1.5x 加重(體檢已扣一次,在 ETF 卡位場景更該避免)
        # 體檢扣 -5/-10/-15 → 這裡扣 -7/-15/-22
        pen = pd.to_numeric(r.get("⑪動態惡化扣分"), errors="coerce")
        if pd.notna(pen) and pen < 0:
            s += float(pen) * 1.5
        ca = pd.to_numeric(r.get("含金量"), errors="coerce")
        if pd.notna(ca):
            if ca >= 1.2: s += 15
            elif ca >= 1.0: s += 10
            elif ca >= 0.8: s += 5
        sig = pd.to_numeric(r.get("改善訊號數"), errors="coerce")
        if pd.notna(sig):
            s += int(sig) * 5
        mo = pd.to_numeric(r.get("最新月營收年增%"), errors="coerce")
        if pd.notna(mo):
            if mo >= 30: s += 15
            elif mo >= 15: s += 10
            elif mo >= 0: s += 3
        # 距離 0050 真實門檻越近越加分(前40=納入線,51-60=候補,跌61=剔除)
        rk = r.get("TWSE排名")
        if pd.notna(rk):
            if rk <= 42: s += 30      # 緊貼前40納入線:再升一階就被動買盤強迫納入
            elif rk <= 50: s += 22    # 距前40一步之遙
            elif rk <= 55: s += 14    # 候補區前段
            elif rk <= 60: s += 8     # 候補區尾(51-60 依市值遞補)
            else: s += 2
        return s

    if len(cand_a):
        cand_a["納入潛力分"] = cand_a.apply(score, axis=1)
        cand_a = cand_a.sort_values(["納入潛力分", "TWSE排名"], ascending=[False, True])

    # ---- 候選 B:blind spot — rank 51-90 但不在我們 PICKS 也不在 0050 ----
    our_codes = set(val["代號"].astype(str)) if not val.empty else set()
    spot = rank[
        (rank["rank"] >= 41) & (rank["rank"] <= 60) &
        (~rank["代號"].isin(KNOWN_0050)) &
        (~rank["代號"].isin(our_codes))
    ].copy()
    spot = spot.rename(columns={"rank": "TWSE排名", "比重%": "大盤比重%"})
    spot["建議"] = "加入 PICKS 抓體檢"

    # ---- 富櫃50 候選 A:我們已體檢 + OTC rank 51-90 ----
    cand_otc_a = pd.DataFrame()
    cand_otc_b = pd.DataFrame()
    if not otc_rank.empty:
        cand_otc_a = base[
            base["OTC排名"].notna() &
            (base["OTC排名"] >= 41) &
            (base["OTC排名"] <= 60)
        ].copy()

        def score_otc(r):
            s = 0
            g = str(r.get("評等", ""))
            if g == "A": s += 30
            elif g == "B": s += 20
            elif g == "C": s += 10
            v = str(r.get("估值", ""))
            if "便宜" in v: s += 25
            elif "合理" in v: s += 18
            elif "偏貴" in v: s += 5
            s += forward_bonus(r)                      # 未來估值(forward PE/PEG)加分
            # ⑪ 動態惡化扣分(1.5x 加重,同 0050)
            pen = pd.to_numeric(r.get("⑪動態惡化扣分"), errors="coerce")
            if pd.notna(pen) and pen < 0:
                s += float(pen) * 1.5
            ca = pd.to_numeric(r.get("含金量"), errors="coerce")
            if pd.notna(ca):
                if ca >= 1.2: s += 15
                elif ca >= 1.0: s += 10
                elif ca >= 0.8: s += 5
            sig = pd.to_numeric(r.get("改善訊號數"), errors="coerce")
            if pd.notna(sig):
                s += int(sig) * 5
            mo = pd.to_numeric(r.get("最新月營收年增%"), errors="coerce")
            if pd.notna(mo):
                if mo >= 30: s += 15
                elif mo >= 15: s += 10
                elif mo >= 0: s += 3
            rk = r.get("OTC排名")
            if pd.notna(rk):
                if rk <= 42: s += 30
                elif rk <= 50: s += 22
                elif rk <= 55: s += 14
                elif rk <= 60: s += 8
                else: s += 2
            return s

        if len(cand_otc_a):
            cand_otc_a["納入潛力分"] = cand_otc_a.apply(score_otc, axis=1)
            cand_otc_a = cand_otc_a.sort_values(["納入潛力分", "OTC排名"], ascending=[False, True])

        # 富櫃50 blind spot
        our = set(val["代號"].astype(str)) if not val.empty else set()
        cand_otc_b = otc_rank[
            (otc_rank["rank"] >= 41) & (otc_rank["rank"] <= 60) &
            (~otc_rank["代號"].isin(our))
        ].copy().rename(columns={"rank": "OTC排名"})
        cand_otc_b["建議"] = "加入 PICKS 抓體檢"

    # ---- OTC 轉上市觀察(上櫃大型股一旦轉上市即直接卡進 0050/TWSE 前段)----
    # 上櫃股不在 0050 選股池;但「轉上市」是已知催化劑(信驊/環球晶等若上市即進前 50)。
    # 依真實規則:市值前20=快速納入(5日內直接入)、前40=納入線、41-60=候補。
    transfer = pd.DataFrame()
    TWSE_R20_YI = float(rank[rank["rank"] == 20]["市值億"].iloc[0]) if "市值億" in rank.columns else None
    TWSE_R40_YI = float(rank[rank["rank"] == 40]["市值億"].iloc[0]) if "市值億" in rank.columns else None
    TWSE_R60_YI = float(rank[rank["rank"] == 60]["市值億"].iloc[0]) if "市值億" in rank.columns else None
    if os.path.exists(OTC) and TWSE_R40_YI:
        otc = pd.read_csv(OTC, dtype={"代號": str})
        big = otc[otc["市值億"] >= TWSE_R60_YI].copy()
        def tier(m):
            if m >= TWSE_R20_YI:  return "🚀轉上市快速納入(前20,5日內直接入)"
            if m >= TWSE_R40_YI:  return "🔥轉上市即進0050(前40納入線)"
            return "⭐轉上市即進候補(41-60)"
        big["轉上市定位"] = big["市值億"].apply(tier)
        big["在我們PICKS"] = big["代號"].apply(lambda c: "✓" if c in (set(val["代號"].astype(str)) if not val.empty else set()) else "")
        transfer = big[["rank", "代號", "名稱", "市值億", "轉上市定位", "在我們PICKS"]].rename(
            columns={"rank": "OTC排名"})

    # ---- 0050 剔除警報:現有成分股逼近/跌出 61 名 = 被動資金即將賣壓(SELL 訊號)----
    # 規則:成分股跌出前 61 名才剔除。rank ≥ 58 即進入危險區,先標記預警。
    evict = rank[(rank["代號"].isin(KNOWN_0050)) & (rank["rank"] >= 58)].copy()
    if not evict.empty:
        def evtier(rk):
            if rk >= 61: return "🔴已跌出61名(剔除壓力)"
            if rk >= 60: return "🟠逼近剔除線(60)"
            return "🟡接近危險區(58-59)"
        evict["剔除警報"] = evict["rank"].apply(evtier)
        evict["在我們PICKS"] = evict["代號"].apply(lambda c: "✓" if c in (set(val["代號"].astype(str)) if not val.empty else set()) else "")
        evict = evict[["rank", "代號", "名稱", "比重%", "剔除警報", "在我們PICKS"]].rename(columns={"rank": "TWSE排名"})

    # ---- 輸出 ----
    out_cols = ["TWSE排名", "代號", "名稱", "大盤比重%", "納入潛力分", "評等", "品質總分",
                "估值", "未來估值", "PEG", "ForwardPE", "成長率g%",
                "⑪動態惡化扣分", "負債比%", "流動比%", "存貨年增%",
                "改善訊號數", "分級", "含金量", "市值(億)", "PER(自算)", "PE位階%",
                "PBR", "殖利率%", "最新月營收年增%", "循環股"]
    out_cols = [c for c in out_cols if c in cand_a.columns]

    # universe 排名表(只顯示有 TWSE 排名的)
    in_twse = base[base["TWSE排名"].notna()].sort_values("TWSE排名")

    os.makedirs("data", exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        cand_a[out_cols].to_excel(xw, sheet_name="候選A_已體檢", index=False)
        spot[["TWSE排名", "代號", "名稱", "大盤比重%", "建議"]].to_excel(
            xw, sheet_name="候選B_blind_spot", index=False)
        in_twse_cols = [c for c in ["TWSE排名", "代號", "名稱", "大盤比重%",
                                     "是否0050", "評等", "估值"] if c in in_twse.columns]
        in_twse[in_twse_cols].to_excel(xw, sheet_name="universe在TWSE排名", index=False)
        rank.head(100).to_excel(xw, sheet_name="TWSE前100", index=False)
        if not cand_otc_a.empty:
            otc_a_cols = ["OTC排名", "代號", "名稱", "OTC市值億", "納入潛力分", "評等",
                          "品質總分", "估值", "未來估值", "PEG", "ForwardPE", "成長率g%",
                          "⑪動態惡化扣分", "負債比%", "流動比%", "存貨年增%",
                          "改善訊號數", "分級", "含金量",
                          "PER(自算)", "PE位階%", "PBR", "殖利率%",
                          "最新月營收年增%", "循環股"]
            otc_a_cols = [c for c in otc_a_cols if c in cand_otc_a.columns]
            cand_otc_a[otc_a_cols].to_excel(xw, sheet_name="富櫃50候選A_已體檢", index=False)
        if not cand_otc_b.empty:
            cand_otc_b[["OTC排名", "代號", "名稱", "市值億", "建議"]].to_excel(
                xw, sheet_name="富櫃50候選B_blind", index=False)
        if not otc_rank.empty:
            otc_rank.head(100).to_excel(xw, sheet_name="OTC前100", index=False)
        if not transfer.empty:
            transfer.to_excel(xw, sheet_name="OTC轉上市觀察", index=False)
        if not evict.empty:
            evict.to_excel(xw, sheet_name="0050剔除警報", index=False)
        thresh_df = pd.DataFrame([
            {"項目": "rank 40 (納入線) 比重%", "值": THRESH_40},
            {"項目": "rank 60 (候補尾) 比重%", "值": THRESH_60},
            {"項目": "rank 61 (剔除線) 比重%", "值": THRESH_61},
            {"項目": "我們 universe 在 TWSE 前 100 內", "值": int((in_twse["TWSE排名"] <= 100).sum())},
            {"項目": "我們 universe 在 rank 41-60 卡位區", "值": int(((in_twse["TWSE排名"] >= 41) & (in_twse["TWSE排名"] <= 60)).sum())},
            {"項目": "候選A (已體檢)", "值": len(cand_a)},
            {"項目": "候選B (blind spot)", "值": len(spot)},
            {"項目": "0050 剔除警報 (成分股逼近61名)", "值": len(evict)},
        ])
        thresh_df.to_excel(xw, sheet_name="門檻說明", index=False)

    print(f"\n完成 → {OUT}")
    print(f"候選A (已體檢) {len(cand_a)} 檔 / 候選B (blind spot) {len(spot)} 檔")
    if len(cand_a):
        show = cand_a.head(15)[[c for c in ["TWSE排名", "代號", "名稱", "納入潛力分",
                                              "評等", "估值", "未來估值", "PEG"]
                                  if c in cand_a.columns]]
        print(f"\n候選A Top 15:\n{show.to_string(index=False)}")
    if len(spot):
        print(f"\n候選B (blind spot) 全部 {len(spot)} 檔:\n"
              f"{spot[['TWSE排名','代號','名稱','大盤比重%']].to_string(index=False)}")
    print(f"\n富櫃50 候選A (已體檢) {len(cand_otc_a)} 檔 / B (blind) {len(cand_otc_b)} 檔")
    if not cand_otc_a.empty:
        show = cand_otc_a.head(10)[[c for c in ["OTC排名","代號","名稱","納入潛力分",
                                                  "評等","估值","未來估值","PEG"]
                                      if c in cand_otc_a.columns]]
        print(f"富櫃50 候選A Top 10:\n{show.to_string(index=False)}")
    if not cand_otc_b.empty:
        print(f"\n富櫃50 候選B (blind) Top 15:\n"
              f"{cand_otc_b.head(15)[['OTC排名','代號','名稱','市值億']].to_string(index=False)}")
    if not transfer.empty:
        print(f"\nOTC 轉上市觀察 {len(transfer)} 檔(轉上市即可進 TWSE 前 60):\n"
              f"{transfer.to_string(index=False)}")
    if not evict.empty:
        print(f"\n🔴 0050 剔除警報 {len(evict)} 檔(成分股逼近 61 名=被動賣壓):\n"
              f"{evict.to_string(index=False)}")
    return cand_a


if __name__ == "__main__":
    main()
