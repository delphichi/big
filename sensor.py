#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SLCA 投資感測器 v2  sensor.py
=============================
把「SLCA 投資感測器 Sensor Prompt v2」的偵測流程程式化。

  感測器的唯一工作:發現差異,產生種子,交棒。它不分析,不下結論,不建議買賣。
  (見 SLCA_______Prompt_v2.md;本檔是該 Prompt 的可執行版)

實作的七步流程(與 Prompt「感測器執行流程」一一對應):
  步驟0  讀北極星 + 能力圈 + 死亡模式庫
  步驟1  硬過濾(市值/流動性/能力圈;死亡模式命中是「扣分」不是排除)
  步驟2  六種差異掃描 + 每個差異算 DS
  步驟3  共振偵測(同標的多種差異 +10/種,上限+30)+ 死亡模式比對(扣分)→ 綜合DS排序
  步驟4  假陽性過濾(波普爾三問;Q1≥3次扣15,Q2平凡解釋成立則不輸出,Q3答不出則不輸出)
  步驟5  機會成本感測(被明顯壓制者降為觀察名單)
  步驟6  信心分數 0–100(基礎 + 來源加分 − 折扣)
  步驟7  依注意力預算輸出(A級DS>85最多1顆、B級DS70–85最多2顆、合計≤3)

設計原則:
  ‧ 引擎只用標準庫(json),沒有 FinMind/pandas 也能跑 → 完全由輸入檔驅動,確保可重現、可測。
  ‧ ①價格/②基本面/③矛盾 這類量化差異,可用 --auto 從 FinMind 自動偵測補進來(選配,需 token)。
  ‧ ④敘事/⑤反共識/⑥時間 與波普爾三問答案,本質是判斷,一律由每週輸入檔提供。

用法:
  python3 sensor.py --input sensor_input.example.json
  python3 sensor.py --input my_week.json --out data/SLCA_種子.md
  python3 sensor.py --input my_week.json --auto      # 額外用 FinMind 自動偵測 ①②③(需 FINMIND_TOKEN)

輸入檔格式見 sensor_input.example.json 與 build_input_template()。
"""

import os
import json
import argparse
import datetime as _dt

# ════════════════════════════════════════════
# 六種差異:基礎DS權重(取 Prompt 區間中位數為預設,可被輸入檔的 base_ds 覆寫)
# ════════════════════════════════════════════
DIFF_TYPES = {
    "價格差異":   dict(rng=(20, 40),  default=30, conf_tag="價格"),
    "基本面差異": dict(rng=(50, 70),  default=60, conf_tag="基本面"),
    "矛盾訊號":   dict(rng=(70, 90),  default=80, conf_tag="矛盾"),
    "敘事差異":   dict(rng=(30, 50),  default=40, conf_tag="敘事"),   # 單獨不成種子;搭配②③加成至60–80
    "反共識裂縫": dict(rng=(80, 95),  default=88, conf_tag="反共識"), # 需具體前提鬆動證據
    "時間差異":   dict(rng=(50, 75),  default=62, conf_tag="時間"),
}
NARRATIVE = "敘事差異"

# ════════════════════════════════════════════
# 死亡模式庫(初始4模式;觸發/歷史/失敗率/處置 對齊 Prompt)
#   penalty:綜合DS扣分(Prompt「-20至-50」);命中後在步驟3扣除。
#   特例 001:額外把「敘事差異」的基礎分砍半(Prompt:敘事DS直接砍半,需基本面矛盾訊號才能回復)。
#   命中由輸入檔宣告 death_patterns: ["002", ...](結構惡化這類事實無法純量化偵測);
#   001 可由 --auto 在「敘事+量暴增+基本面無改善」時自動標記。
# ════════════════════════════════════════════
DEATH_PATTERNS = {
    "001": dict(name="敘事泡沫", penalty=40,
                trigger="敘事爆發 + 成交量暴增 + 基本面無對應改善",
                history="NFT(2021)、SPAC(2021)、元宇宙(2022)、AI概念股初期",
                fail_rate="~85%", handling="敘事DS直接砍半,需基本面矛盾訊號才能回復"),
    "002": dict(name="假底部", penalty=35,
                trigger="股價創低 + 法人買超 + 但產業結構惡化中",
                history="傳統零售、面板(週期底部誤判)",
                fail_rate="~70%", handling="加入『產業結構檢驗』為必要條件"),
    "003": dict(name="業績轉機幻覺", penalty=30,
                trigger="單季獲利大幅改善 + 市場解讀為轉機",
                history="多數景氣循環股底部反彈後再破底",
                fail_rate="~65%", handling="要求連續兩季改善才計入DS,單季不算"),
    "004": dict(name="政策題材", penalty=25,
                trigger="政府政策宣布 + 相關股大漲",
                history="各國新能源補貼概念股(多數最終回落)",
                fail_rate="~60%", handling="需『政策落地時程明確』+『公司已有營收受益』才允許進DS"),
}

# 注意力預算
A_MIN = 85          # A級:DS > 85
B_MIN = 70          # B級:DS 70–85
A_MAX, B_MAX = 1, 2 # A最多1、B最多2,合計≤3
WATCH_LO, WATCH_HI = 50, 70   # 觀察名單區間


# ════════════════════════════════════════════
# 步驟0:載入輸入 + 死亡模式庫(可被輸入檔擴充)
# ════════════════════════════════════════════
def load_input(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # 合併使用者自訂死亡模式(Prompt:死亡模式庫要更新)
    for p in data.get("death_patterns_extra", []):
        DEATH_PATTERNS[str(p["id"])] = dict(
            name=p.get("name", ""), penalty=int(p.get("penalty", 30)),
            trigger=p.get("trigger", ""), history=p.get("history", ""),
            fail_rate=p.get("fail_rate", ""), handling=p.get("handling", ""))
    return data


# ════════════════════════════════════════════
# 步驟1:硬過濾(不符北極星/能力圈/市值/流動性 → 直接剔除;死亡模式不在此排除)
# ════════════════════════════════════════════
def hard_filter(tk, north):
    cap_min = north.get("market_cap_min_yi")
    reasons = []
    if tk.get("in_circle") is False:
        reasons.append("不在能力圈")
    if tk.get("liquidity_ok") is False:
        reasons.append("流動性不足")
    cap = tk.get("market_cap_yi")
    if cap_min is not None and cap is not None and cap < cap_min:
        reasons.append(f"市值{cap}億 < 門檻{cap_min}億")
    return reasons   # 空 = 通過


# ════════════════════════════════════════════
# 步驟2+3:六種差異掃描 → 每個差異 DS → 共振 + 死亡模式 → 綜合DS
# ════════════════════════════════════════════
def base_of(diff, types_present):
    """單一差異的基礎分。敘事搭配②③時提升;敘事單獨保持低分(且鐵律3會擋掉輸出)。"""
    t = diff["type"]
    base = diff.get("base_ds")            # 缺鍵或填 null 都回退到該差異類型的預設權重
    if base is None:
        base = DIFF_TYPES[t]["default"]
    if t == NARRATIVE:
        combined = ("基本面差異" in types_present) or ("矛盾訊號" in types_present)
        if combined:
            base = max(base, 70)          # 搭配②③ → 加成至 60–80 區間
    return float(base)


def score_ds(tk, north):
    """回傳 (綜合DS, calc 明細 dict)。"""
    diffs = tk.get("differences", [])
    types_present = [d["type"] for d in diffs]
    distinct = list(dict.fromkeys(types_present))

    # 敘事單獨 → 鐵律3:不成種子(標記,綜合DS視為無效)
    narrative_only = (set(distinct) == {NARRATIVE})

    # 死亡模式001:敘事基礎分砍半(在算 base 前處理)
    hit = [str(p) for p in tk.get("death_patterns", [])]
    halve_narrative = "001" in hit

    bases = []
    for d in diffs:
        b = base_of(d, types_present)
        if d["type"] == NARRATIVE and halve_narrative:
            b /= 2.0
        bases.append((d["type"], b))
    base_score = max((b for _, b in bases), default=0.0)
    primary = max(bases, key=lambda x: x[1])[0] if bases else None

    # 共振加分:同標的多種「差異類型」,每多一種 +10,上限 +30
    resonance = min(max(len(distinct) - 1, 0) * 10, 30)

    # 歷史加分:該(主)差異類型在北極星下過去有成功案例(輸入檔 history_success 提供 +5~15)
    hist_map = north.get("history_success", {}) if isinstance(north, dict) else {}
    history = 0
    if primary and primary in hist_map:
        history = max(5, min(15, int(hist_map[primary])))

    # 死亡模式扣分(每個命中模式扣其 penalty,單模式限 -20~-50)
    death = 0
    death_detail = []
    for pid in hit:
        if pid in DEATH_PATTERNS:
            pen = max(20, min(50, DEATH_PATTERNS[pid]["penalty"]))
            death += pen
            death_detail.append((pid, DEATH_PATTERNS[pid]["name"], pen))

    ds = base_score + resonance + history - death
    if narrative_only:
        ds = 0.0   # 鐵律3
    ds = round(max(0.0, min(100.0, ds)), 1)

    calc = dict(types=distinct, primary=primary, base=round(base_score, 1),
                resonance=resonance, history=history, death=death,
                death_detail=death_detail, narrative_only=narrative_only,
                halve_narrative=halve_narrative)
    return ds, calc


# ════════════════════════════════════════════
# 步驟4:波普爾三問(只對 DS>70 候選做)
#   Q1 ≥3 → DS-15 並記為潛在假陽性;Q2 平凡解釋成立 → 不輸出;Q3 答不出 → 不輸出
# ════════════════════════════════════════════
def popper_filter(ds, popper):
    p = popper or {}
    q1 = int(p.get("q1_false_positive_count", 0))
    q2_fully = bool(p.get("q2_fully_explains", False))
    q3 = (p.get("q3_observable") or "").strip()

    adj = 0
    notes = dict()
    if q1 >= 3:
        adj -= 15
        notes["q1"] = f"出現{q1}次未帶來機會 → DS-15(記為潛在假陽性模式)"
    else:
        notes["q1"] = f"出現{q1}次,未達3次門檻"

    if q2_fully:
        notes["q2"] = "平凡解釋已能完全解釋差異 → 不輸出(記為已解釋異常)"
        return False, adj, notes
    notes["q2"] = (p.get("q2_mundane") or "—") + "(不足以完全解釋差異)"

    if not q3:
        notes["q3"] = "答不出12個月內可觀察事件 → 不輸出(故事,非假說)"
        return False, adj, notes
    notes["q3"] = q3
    return True, adj, notes


# ════════════════════════════════════════════
# 步驟6:信心分數 0–100
# ════════════════════════════════════════════
def confidence(tk, calc):
    types = set(calc["types"])
    has_premise = bool((tk.get("fragile_premise") or "").strip())

    # 信心基礎(取最符合的組合)
    if "反共識裂縫" in types and has_premise:
        base = 90
    elif "基本面差異" in types and "矛盾訊號" in types:
        base = 80
    elif "價格差異" in types and "基本面差異" in types:
        base = 60
    elif types == {"價格差異"}:
        base = 35
    else:
        # Prompt 只列4種組合;其餘以主差異基礎DS粗估,保守落在合理區間
        base = int(min(80, max(35, calc["base"])))

    bonus = 0
    src = []
    if tk.get("has_history_case"):
        bonus += 5; src.append("歷史案例")
    if any(d.get("institutional_support") for d in tk.get("differences", [])):
        bonus += 5; src.append("法人動向")
    if len(types) >= 3:
        bonus += 10; src.append("多重共振")

    disc = 0
    if calc["primary"] == NARRATIVE:
        disc -= 10
    if tk.get("market_volatility_high") or tk.get("_vol_high"):
        disc -= 5
    if tk.get("kill_hard_to_verify"):
        disc -= 10

    score = max(0, min(100, base + bonus + disc))
    return score, dict(base=base, bonus=bonus, disc=disc, src=src)


# ════════════════════════════════════════════
# 組裝:跑完步驟0–6,回傳候選/觀察/忽略/淘汰
# ════════════════════════════════════════════
def run_sensor(data):
    north = data.get("north_star", {})
    vol_high = bool(data.get("market_volatility_high", False))

    candidates, watch, ignored, killed = [], [], [], []
    for tk in data.get("tickers", []):
        tk["_vol_high"] = vol_high
        label = f'{tk.get("id","?")} {tk.get("name","")}'.strip()

        # 步驟1 硬過濾
        fr = hard_filter(tk, north)
        if fr:
            killed.append((label, "硬過濾:" + "、".join(fr)))
            continue

        # 步驟2+3 DS
        ds, calc = score_ds(tk, north)
        rec = dict(label=label, tk=tk, ds=ds, calc=calc)
        why = _death_note(calc)

        if calc["narrative_only"]:
            killed.append((label, "鐵律3:敘事單獨不成種子"))
            continue
        if ds < WATCH_LO:
            ignored.append((label, ds, why)); continue
        if ds < WATCH_HI:
            watch.append((label, ds, why)); continue
        candidates.append(rec)   # DS>70 → 候選池

    # 綜合DS排序,取前3名進入假陽性過濾(Prompt)
    candidates.sort(key=lambda r: r["ds"], reverse=True)
    pool, overflow = candidates[:3], candidates[3:]
    for r in overflow:
        watch.append((r["label"], r["ds"], "綜合DS未進前3,本週先觀察"))

    # 步驟4 波普爾三問
    survivors = []
    for r in pool:
        ok, adj, notes = popper_filter(r["ds"], r["tk"].get("popper"))
        r["ds"] = round(max(0.0, min(100.0, r["ds"] + adj)), 1)
        r["popper_notes"] = notes
        if not ok:
            killed.append((r["label"], "波普爾三問未過:" +
                           (notes.get("q2") if "已解釋異常" in notes.get("q2", "")
                            else notes.get("q3", ""))))
            continue
        if r["ds"] < WATCH_HI:     # 三問扣分後跌破70 → 降觀察
            watch.append((r["label"], r["ds"], "波普爾Q1扣分後跌破70"))
            continue
        survivors.append(r)

    # 步驟6 信心分數(機會成本前先算,供排序與壓制判斷)
    for r in survivors:
        r["conf"], r["conf_calc"] = confidence(r["tk"], r["calc"])

    # 步驟5 機會成本 + 步驟7 注意力預算
    survivors.sort(key=lambda r: (r["ds"], r["conf"]), reverse=True)
    seeds, demoted = allocate_budget(survivors)
    for r in demoted:
        watch.append((r["label"], r["ds"], r.get("opp_cost", "機會成本:被壓制")))

    return dict(seeds=seeds, watch=watch, ignored=ignored, killed=killed,
                north=north, as_of=data.get("as_of"))


def _death_note(calc):
    if not calc["death_detail"]:
        return ""
    return "命中死亡模式 " + "、".join(
        f"{pid} {name}(-{pen})" for pid, name, pen in calc["death_detail"])


def allocate_budget(survivors):
    """步驟5+7:A級≤1、B級≤2、合計≤3。塞不下的視為被壓制 → 降觀察(機會成本)。"""
    seeds, demoted = [], []
    a_used = b_used = 0
    for i, r in enumerate(survivors):
        grade = "A" if r["ds"] > A_MIN else "B"
        placed = False
        if grade == "A" and a_used < A_MAX:
            a_used += 1; placed = True
        elif grade == "B" and b_used < B_MAX:
            b_used += 1; placed = True
        # A級滿了但還有B級空間 → A級無法降級成B(分級稀缺),只能壓制
        if placed and (a_used + b_used) <= (A_MAX + B_MAX):
            r["grade"] = grade
            r["rank"] = i + 1
            r["opp_cost"] = f"本週候選第{i+1}名,獲配{grade}級額度"
            seeds.append(r)
        else:
            top = survivors[0]
            r["opp_cost"] = (f"本週候選第{i+1}名,{grade}級額度已滿,"
                             f"被 {top['label']}(DS{top['ds']}) 壓制 → 降觀察名單")
            demoted.append(r)
    return seeds, demoted


# ════════════════════════════════════════════
# 輸出規格(每顆種子格式,對齊 Prompt「輸出規格」)
# ════════════════════════════════════════════
def render_seed(n, r):
    tk = r["tk"]; calc = r["calc"]; notes = r["popper_notes"]; cc = r["conf_calc"]
    dd = "、".join(f"{pid} {name}(-{pen})" for pid, name, pen in calc["death_detail"]) or "未命中"
    src = "／".join([DIFF_TYPES[t]["conf_tag"] for t in calc["types"]])
    conf_src = ("、".join(cc["src"]) or "—")
    L = []
    L.append(f"種子 #{n}  [{r['grade']}級]")
    L.append("─" * 32)
    L.append(f"標的:{r['label']}")
    L.append(f"差異類型:{ '＋'.join(calc['types']) }")
    L.append(f"綜合DS:{r['ds']}")
    L.append(f"信心分數:{r['conf']}")
    L.append(f"信心來源:{src}（加分:{conf_src}）")
    L.append("")
    L.append(f"觸發原因:{tk.get('trigger') or _auto_trigger(tk)}")
    L.append(f"市場共識:{tk.get('consensus') or '—'}")
    L.append(f"共識的脆弱前提:{tk.get('fragile_premise') or '—'}")
    L.append("")
    L.append("波普爾三問回答:")
    L.append(f"  Q1 歷史假陽性:{notes.get('q1','—')}")
    L.append(f"  Q2 平凡解釋:{notes.get('q2','—')}")
    L.append(f"  Q3 可觀察事件:{notes.get('q3','—')}")
    L.append("")
    L.append(f"死亡模式比對:{dd}")
    L.append(f"機會成本:{r.get('opp_cost','—')}")
    L.append("")
    L.append(f"建議SLCA方向:{tk.get('suggest_direction') or '深化'}")
    L.append(f"殺死條件:{tk.get('kill_condition') or '—'}")
    L.append(f"觀察期:{tk.get('observe_months','—')} 個月")
    L.append(f"緊急程度:{_urgency(r)}")
    L.append("─" * 32)
    return "\n".join(L)


def _auto_trigger(tk):
    return "；".join(d.get("note", d["type"]) for d in tk.get("differences", [])) or "—"


def _urgency(r):
    if r["grade"] == "A" and r["conf"] >= 80:
        return "高"
    if r["ds"] >= 80 or r["conf"] >= 75:
        return "中"
    return "低"


def render_report(res, full=False):
    L = []
    L.append(f"# SLCA 投資感測器 v2 · 本週種子")
    L.append(f"> 掃描日期:{res.get('as_of') or _dt.date.today().isoformat()}　"
             f"｜　輸出 {len(res['seeds'])} 顆（注意力預算:A≤1 / B≤2 / 合計≤3）")
    n = res["north"]
    if n:
        L.append("")
        L.append(f"**北極星**:{n.get('philosophy','—')}　"
                 f"｜ 能力圈:{('、'.join(n.get('circle', [])) or '—')}　"
                 f"｜ 市值門檻:{n.get('market_cap_min_yi','—')} 億　"
                 f"｜ 時間框架:{n.get('timeframe','—')}")
    L.append("")
    if res["seeds"]:
        L.append("```")
        for i, r in enumerate(res["seeds"], 1):
            L.append(render_seed(i, r))
            L.append("")
        L.append("```")
    else:
        L.append("> **空週**。本週無差異通過全部過濾 —— 沒有差異就是沒有機會,不發明機會(鐵律2)。")

    def fmt(items, cap=20):
        if not items:
            return "—"
        items = sorted(items, key=lambda it: it[1], reverse=True)   # 依DS高到低
        out = []
        for it in items[:cap]:
            lab, ds = it[0], it[1]
            why = it[2] if len(it) > 2 else ""
            out.append(f"{lab}(DS{ds}{('；' + why) if why else ''})")
        lines = ["    - " + x for x in out]
        if len(items) > cap:
            lines.append(f"    - …其餘 {len(items) - cap} 檔(DS 較低,略)")
        return "\n".join(lines)
    L.append("")
    L.append("---")
    L.append("## 其餘標的去向(透明化)")
    L.append("- **觀察名單(DS 50–70 或被壓制)**:")
    L.append(fmt(res['watch']))
    L.append("- **忽略(DS < 50)**:")
    L.append(fmt(res['ignored']))
    if res["killed"]:
        L.append("- **過濾/淘汰**:")
        for lab, why in res["killed"]:
            L.append(f"    - {lab}:{why}")
    L.append("")
    L.append("---")
    L.append("## 交棒 SLCA v5")
    if res["seeds"]:
        L.append("複製種子全文(含DS、信心分數、三問回答),貼入 SLCA v5 系統Prompt,")
        L.append("並依各顆「信心分數」決定深化力度;A級才值得全套深化分析。")
    else:
        L.append("本週無種子可交棒。")

    if full and res["seeds"]:
        verify = res.get("verify")
        L.append("")
        L.append(render_handoff(res))
        L.append("")
        L.append(render_reality_check(res, verify))
        L.append("")
        L.append(render_evolution_log(res, verify))
    return "\n".join(L)


def render_handoff(res):
    """步驟:交棒 SLCA v5 —— 每顆種子產生一段可直接貼上的指令(對齊 Prompt『與 SLCA v5 的銜接』)。"""
    L = ["---", "## 交棒指令(逐顆,直接複製貼入 SLCA v5 對話)"]
    for i, r in enumerate(res["seeds"], 1):
        tk = r["tk"]
        L.append("")
        L.append(f"### 種子 #{i}　{r['label']}　[{r['grade']}級]")
        L.append("```")
        L.append(render_seed(i, r))
        L.append("")
        L.append("以這顆種子為起點執行 SLCA 認知循環。")
        L.append(f"建議方向:{tk.get('suggest_direction') or '深化'}")
        L.append(f"信心分數 {r['conf']},請依此決定深化力度。")
        L.append("請特別針對殺死條件與共識脆弱前提執行紅隊攻擊。")
        L.append("```")
    return "\n".join(L)


def render_reality_check(res, verify=None):
    """步驟:現實驗證 —— Q3 可觀察事件 + 殺死條件 + 檢核截止;
    若有 verify(total_screener 抓回的客觀數據),附上現況快照與自動判讀。"""
    verify = verify or {}
    as_of = res.get("as_of") or _dt.date.today().isoformat()
    try:
        base = _dt.date.fromisoformat(as_of)
    except Exception:
        base = _dt.date.today()
    L = ["---", "## 現實驗證追蹤表(Reality Loop · Q3 是否兌現)"]
    for i, r in enumerate(res["seeds"], 1):
        tk = r["tk"]
        months = int(tk.get("observe_months") or 12)
        due = base + _dt.timedelta(days=months * 30)   # 近似:每月30天,免依賴外部套件
        q3 = (tk.get("popper", {}) or {}).get("q3_observable") or "—"
        kill = tk.get("kill_condition") or "—"
        sid = tk.get("id")
        L.append("")
        L.append(f"### #{i} {r['label']}　差異:{'＋'.join(r['calc']['types'])}")
        L.append(f"- **可觀察事件(Q3)**:{q3}")
        L.append(f"- **殺死條件**:{kill}")
        L.append(f"- **檢核截止**:{due.isoformat()}(觀察期 {months} 個月)")
        snap = verify.get(sid)
        if snap is None:
            L.append("- **客觀現況**:(未取得 FinMind 數據;設定 FINMIND_TOKEN 後以 --full 重跑即自動填入)")
            L.append("- **狀態**:☐ 待驗證")
        else:
            L.append("- **客觀現況(total_screener 抓回)**:")
            for line in _snap_lines(snap):
                L.append(f"    - {line}")
            L.append(f"- **自動判讀**:{_verdict(r['calc']['types'], snap)}")
            L.append("- **狀態**:☐ 待人工確認(數據已附,Q3 文字事件仍需人眼判定)")
    L.append("")
    L.append("> 驗證後:兌現 → 強化該差異類型 DS 權重;殺死條件成立 → 記錄並分析哪步出錯;")
    L.append("> 假陽性反覆出現 → 考慮新增死亡模式。結果回填下方演化記錄,並回寫感測器。")
    return "\n".join(L)


def render_evolution_log(res, verify=None):
    """步驟:演化記錄欄 —— 本次掃描的紀錄列;若有 verify,把客觀數據快照寫進『掃描當下數據』。"""
    verify = verify or {}
    as_of = res.get("as_of") or _dt.date.today().isoformat()
    n = len(res["seeds"])
    types = "、".join(sorted({"＋".join(r["calc"]["types"]) for r in res["seeds"]})) or "—"
    confs = "／".join(str(r["conf"]) for r in res["seeds"]) or "—"
    dss = "／".join(str(r["ds"]) for r in res["seeds"]) or "—"
    L = ["---", "## 演化記錄欄(本次掃描;結果欄待現實驗證後回填)",
         "", "```",
         "[日期] | 輸出N顆 | 差異類型 | DS | 信心分數 | 後來結果 | 學到什麼 | 死亡模式庫變動",
         f"{as_of} | {n}顆 | {types} | {dss} | {confs} | (待驗證) | (待回填) | (無)",
         "```"]
    if verify:
        L.append("")
        L.append("### 掃描當下客觀數據快照(total_screener,供日後比對差異是否收斂)")
        L.append("")
        L.append("| 種子 | 五維分類 | 總評分 | 共振 | ROE百分位 | 近四季FCF(億) | 毛利率趨勢 | 連兩季淨利改善 | 股價位置 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for i, r in enumerate(res["seeds"], 1):
            s = verify.get(r["tk"].get("id")) or {}
            L.append("| #{} {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
                i, r["label"], s.get("分類", "—"), s.get("總評分", "—"),
                s.get("共振分數", "—"), s.get("ROE百分位", "—"), s.get("近四季FCF", "—"),
                s.get("毛利率趨勢", "—"), ("是" if s.get("連兩季淨利改善") else
                                        ("否" if s.get("連兩季淨利改善") is not None else "—")),
                s.get("股價位置", "—")))
    L.append("")
    L.append("> 將此列貼到專案的演化記錄;待 Q3 事件或殺死條件揭曉後回填『後來結果/學到什麼』,")
    L.append("> 並依學習更新 DS 權重與死亡模式庫(Reality Loop 的回寫)。")
    return "\n".join(L)


# ════════════════════════════════════════════
# 現實驗證:用 total_screener 抓 FinMind 數據,為每顆種子算客觀指標
#   沒裝 FinMind / 沒 token 時回傳空 dict,render 端自動退回「待驗證」。
# ════════════════════════════════════════════
def verify_seeds(seeds):
    if not seeds:
        return {}
    try:
        import pandas as pd
        from total_screener import (make_loader, fetch_all, fetch_price, assess,
                                     pivot, pick, BENCHMARK, START_PRICE)
    except Exception as e:
        print(f"[現實驗證 跳過] 無法載入 FinMind / total_screener:{e}")
        return {}
    if not os.environ.get("FINMIND_TOKEN"):
        print("[現實驗證 跳過] 未設定 FINMIND_TOKEN(資料量大,務必設)")
        return {}

    dl = make_loader()
    bench = fetch_price(dl, BENCHMARK, START_PRICE)
    bench = bench["close"] if not bench.empty else pd.Series(dtype=float)
    out = {}
    for r in seeds:
        sid = r["tk"].get("id")
        if not sid:
            continue
        try:
            raw = fetch_all(dl, sid)
            rr, _, _ = assess(sid, raw, bench)          # 沿用五維篩選的所有指標
            snap = {k: rr.get(k) for k in
                    ("分類", "總評分", "共振分數", "五關", "PE", "PE百分位",
                     "ROE%", "ROE百分位", "近四季FCF", "含金量",
                     "相對報酬%", "週斜率%", "RS創高", "亮燈")}
            snap.update(_margin_ni_price(raw, pivot, pick))   # 補:毛利率趨勢/連兩季淨利/股價位置
            out[sid] = snap
            print(f"[現實驗證] {sid} 已抓取並分析")
        except Exception as e:
            print(f"[現實驗證] {sid} 失敗:{e}")
    return out


def _margin_ni_price(raw, pivot, pick):
    """補算 total_screener 沒直接給的三項:近四季毛利率趨勢、連兩季淨利是否改善、股價在52週區間位置。"""
    import pandas as pd
    snap = {"毛利率近4季": None, "毛利率趨勢": None, "連兩季淨利改善": None, "股價位置": None}
    inc = pivot(raw.get("inc"))
    if not inc.empty:
        rev = pd.to_numeric(pick(inc, "Revenue", "OperatingRevenue", "NetSales"), errors="coerce")
        gp = pd.to_numeric(pick(inc, "GrossProfit", "GrossProfitLoss"), errors="coerce")
        gm = (gp / rev * 100).dropna()
        if len(gm) >= 2:
            last4 = [round(x, 1) for x in gm.tail(4).tolist()]
            snap["毛利率近4季"] = last4
            snap["毛利率趨勢"] = ("上升" if gm.iloc[-1] > gm.iloc[-2] else
                               "下滑" if gm.iloc[-1] < gm.iloc[-2] else "持平")
        ni = pd.to_numeric(pick(inc, "IncomeAfterTaxes", "IncomeAfterTax", "ProfitAfterTax"),
                           errors="coerce").dropna()
        if len(ni) >= 3:
            snap["連兩季淨利改善"] = bool(ni.iloc[-1] > ni.iloc[-2] and ni.iloc[-2] > ni.iloc[-3])
    price = raw.get("price")
    if price is not None and not price.empty and len(price) >= 252:
        s = price["close"]; last = s.iloc[-1]
        lo, hi = s.tail(252).min(), s.tail(252).max()
        if hi > lo:
            pos = (last - lo) / (hi - lo)
            snap["股價位置"] = ("接近52週低" if pos < 0.2 else
                             "接近52週高" if pos > 0.8 else f"區間{round(pos*100)}%位置")
    return snap


def _snap_lines(s):
    """把快照整理成可讀的條列(現實驗證用)。"""
    L = []
    L.append(f"五維分類:{s.get('分類','—')}　總評分 {s.get('總評分','—')}　共振 {s.get('共振分數','—')}　五關 {s.get('五關','—')}")
    L.append(f"估值/品質:PE {s.get('PE','—')}(歷史百分位 {s.get('PE百分位','—')})、"
             f"ROE {s.get('ROE%','—')}%(歷史百分位 {s.get('ROE百分位','—')})、"
             f"近四季FCF {s.get('近四季FCF','—')} 億、含金量 {s.get('含金量','—')}")
    L.append(f"毛利率近4季:{s.get('毛利率近4季','—')}　趨勢:{s.get('毛利率趨勢','—')}　"
             f"連兩季淨利改善:{'是' if s.get('連兩季淨利改善') else ('否' if s.get('連兩季淨利改善') is not None else '—')}")
    L.append(f"動能:股價位置 {s.get('股價位置','—')}、近半年相對報酬 {s.get('相對報酬%','—')}%、"
             f"週斜率 {s.get('週斜率%','—')}%/週、RS創高 {s.get('RS創高') or '—'}")
    return L


def _verdict(types, s):
    """依種子的差異類型 + 客觀快照,給輕量啟發式判讀(仍需人眼確認)。"""
    types = set(types); notes = []
    roe_p = s.get("ROE百分位"); pos = s.get("股價位置") or ""; gm_t = s.get("毛利率趨勢")
    ni2 = s.get("連兩季淨利改善"); rel = s.get("相對報酬%")
    cashq = s.get("含金量"); fcf = s.get("近四季FCF")
    if "矛盾訊號" in types:
        # 品質/現金背離:含金量<0.8 或 FCF<0 → 品質可能是帳面幻覺
        if (cashq is not None and cashq < 0.8) or (fcf is not None and fcf < 0):
            notes.append(f"含金量{cashq}、近四季FCF{fcf}億 仍偏弱 → 品質/現金背離未解,留意品質幻覺/燒錢")
        # 落後型矛盾:看相對報酬而非絕對股價位置
        elif rel is not None and rel < 0 and roe_p is not None and roe_p >= 70:
            notes.append(f"ROE百分位{roe_p}(品質高)但近半年相對報酬{rel}%(仍落後大盤) → 矛盾未收斂")
        elif rel is not None and rel > 0:
            notes.append(f"相對報酬已轉正({rel}%) → 矛盾收斂中,留意是否已被定價")
        elif roe_p is not None and roe_p >= 70 and "低" in pos:
            notes.append(f"ROE仍居高位(百分位{roe_p})而股價仍偏低 → 品質/定價背離尚在")
    if "基本面差異" in types and gm_t:
        notes.append(f"毛利率趨勢{gm_t}" + ("(支持基本面假說)" if gm_t == "上升"
                     else "(與基本面假說相左)" if gm_t == "下滑" else ""))
    if "價格差異" in types and pos:
        notes.append(f"股價{pos}")
    if ni2 is not None:
        notes.append("連兩季淨利改善" + ("成立(死亡模式003『單季幻覺』風險降低)" if ni2
                     else "不成立(僅單季,留意業績轉機幻覺)"))
    if types & {"時間差異", "敘事差異", "反共識裂縫"}:
        notes.append("事件/敘事面(時程、共識前提等)非財報可量化,需人工追蹤")
    return "；".join(notes) if notes else "客觀數據已附,請人工對照 Q3 與殺死條件判定"


# ════════════════════════════════════════════
# 選配:--auto 用 FinMind 自動偵測 ①價格 / ②基本面 / ③矛盾,合併進輸入檔的標的
#   (沒裝 FinMind 或沒 token 時,自動跳過,不影響引擎)
# ════════════════════════════════════════════
def auto_detect(data):
    try:
        import numpy as np
        import pandas as pd
        from total_screener import (make_loader, fetch_all, fetch_price, momentum,
                                     gate_growth, gate_cash, roe_roic_series, pctile,
                                     BENCHMARK, START_PRICE)
    except Exception as e:
        print(f"[--auto 跳過] 無法載入 FinMind / total_screener:{e}")
        return data
    if not os.environ.get("FINMIND_TOKEN"):
        print("[--auto 跳過] 未設定 FINMIND_TOKEN(資料量大,務必設)")
        return data

    dl = make_loader()
    bench = fetch_price(dl, BENCHMARK, START_PRICE)
    bench = bench["close"] if not bench.empty else pd.Series(dtype=float)
    for tk in data.get("tickers", []):
        sid = tk.get("id")
        if not sid:
            continue
        try:
            raw = fetch_all(dl, sid)
        except Exception as e:
            print(f"[--auto] {sid} 取數失敗:{e}"); continue
        auto_diffs = []

        # ① 價格差異:52週新低但量縮 / 新高但量未放大
        price = raw.get("price")
        if price is not None and not price.empty and len(price) >= 252:
            s, v = price["close"], price["vol"]
            last = s.iloc[-1]
            yr_lo, yr_hi = s.tail(252).min(), s.tail(252).max()
            vol_shrink = v.tail(20).mean() < v.tail(60).mean()
            if last <= yr_lo * 1.03 and vol_shrink:
                auto_diffs.append(dict(type="價格差異",
                                       note="逼近52週新低但成交量萎縮(無人拋售,只是沒人買)"))
            if last >= yr_hi * 0.97 and not (v.tail(20).mean() > v.tail(60).mean() * 1.5):
                auto_diffs.append(dict(type="價格差異",
                                       note="逼近52週新高但量未爆量(強勢但不瘋狂)"))

        # ② 基本面差異:營收持續正成長 / 品質(ROE)歷史高百分位
        pos, win = gate_growth(raw)
        if pos is not None and win and pos >= 10:
            auto_diffs.append(dict(type="基本面差異",
                                   note=f"近{win}月有{pos}月營收YoY正成長(動能持續)"))
        q = roe_roic_series(raw)
        roe_p = None
        if not q.empty and q["ROE"].notna().any():
            roe = q["ROE"].dropna().iloc[-1]
            roe_p = pctile(q["ROE"], roe)
            if roe_p is not None and roe_p >= 80:
                auto_diffs.append(dict(type="基本面差異",
                                       note=f"近四季ROE居自身歷史{roe_p}百分位(品質創高)"))

        # ③ 矛盾訊號:ROE創高(品質強) 但 股價逼近52週低(市場給低分)
        if (price is not None and not price.empty and len(price) >= 252
                and roe_p is not None and roe_p >= 80
                and price["close"].iloc[-1] <= price["close"].tail(252).min() * 1.1):
            auto_diffs.append(dict(type="矛盾訊號",
                                   note="ROE創新高,股價卻接近52週低(品質與定價背離)"))

        if auto_diffs:
            existing = {d["type"] for d in tk.get("differences", [])}
            tk.setdefault("differences", [])
            for d in auto_diffs:
                if d["type"] not in existing:   # 不蓋掉輸入檔已宣告的差異
                    tk["differences"].append(d)
                    existing.add(d["type"])
            print(f"[--auto] {sid} 自動補入差異:{[d['type'] for d in auto_diffs]}")
    return data


# ════════════════════════════════════════════
# 輸入模板(供 --template 產生空白週報輸入)
# ════════════════════════════════════════════
def build_input_template():
    return {
        "as_of": _dt.date.today().isoformat(),
        "north_star": {
            "philosophy": "（你的投資哲學)",
            "circle": ["（能力圈產業1)", "（能力圈產業2)"],
            "market_cap_min_yi": 100,
            "timeframe": "12–24個月",
            "history_success": {"矛盾訊號": 10, "反共識裂縫": 8}
        },
        "market_volatility_high": False,
        "death_patterns_extra": [],
        "tickers": [{
            "id": "XXXX", "name": "（名稱)",
            "market_cap_yi": 0, "in_circle": True, "liquidity_ok": True,
            "differences": [{
                "type": "矛盾訊號", "base_ds": None,
                "note": "（哪裡不對勁)", "institutional_support": False
            }],
            "consensus": "（市場目前怎麼看,一句話)",
            "fragile_premise": "（這件事若不成立,共識就崩潰)",
            "death_patterns": [],
            "has_history_case": False,
            "kill_condition": "（哪個事實成立,這顆種子當場報廢)",
            "kill_hard_to_verify": False,
            "observe_months": 12,
            "suggest_direction": "深化",
            "trigger": "",
            "popper": {
                "q1_false_positive_count": 0,
                "q2_mundane": "（最簡單的平凡解釋)",
                "q2_fully_explains": False,
                "q3_observable": "（12個月內哪個具體事件會驗證或推翻)"
            }
        }]
    }


def main():
    ap = argparse.ArgumentParser(description="SLCA 投資感測器 v2")
    ap.add_argument("--input", help="每週輸入 JSON 檔路徑")
    ap.add_argument("--out", default="data/SLCA_種子.md", help="種子報告輸出路徑(.md)")
    ap.add_argument("--auto", action="store_true", help="額外用 FinMind 自動偵測①②③(需 FINMIND_TOKEN)")
    ap.add_argument("--template", action="store_true", help="印出空白輸入模板 JSON 後結束")
    ap.add_argument("--full", action="store_true",
                    help="種子後附完整 Reality Loop:交棒指令 + 現實驗證追蹤表 + 演化記錄列")
    args = ap.parse_args()

    if args.template or not args.input:
        if not args.input:   # 提示走 stderr,避免污染被重導向的 JSON(stdout)
            import sys
            print("（未提供 --input,印出空白輸入模板;請填好後以 --input 執行)", file=sys.stderr)
        print(json.dumps(build_input_template(), ensure_ascii=False, indent=2))
        return

    data = load_input(args.input)
    if args.auto:
        data = auto_detect(data)

    res = run_sensor(data)
    if args.full:   # 現實驗證:用 total_screener 抓 FinMind 數據(無 token 自動跳過)
        res["verify"] = verify_seeds(res["seeds"])
    report = render_report(res, full=args.full)

    if args.out:
        os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(report + "\n")
        print(f"已輸出種子報告:{args.out}\n")
    print(report)


if __name__ == "__main__":
    main()
