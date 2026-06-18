# 🏦 台股千張大戶週報自動爬蟲

每週五收盤後自動抓取聚財網集保股權分散表，追蹤千張大戶動向。

## 📁 檔案結構

```
├── .github/workflows/
│   └── scrape_holders.yml   ← GitHub Actions 定時任務
├── scripts/
│   └── scraper.py           ← 主爬蟲程式
├── data/                    ← 輸出資料（自動 commit）
│   ├── holders_all.csv      ← 完整歷史資料（可 git diff 追蹤）
│   └── 千張大戶_YYYYMMDD.xlsx  ← 每次執行產出的 Excel
├── requirements.txt
└── README.md
```

## 🚀 快速開始（3步驟）

### 步驟一：Fork 這個 repo

點右上角 **Fork** → 複製到自己的帳號

### 步驟二：啟用 Actions 寫入權限

`Settings` → `Actions` → `General`
→ `Workflow permissions` → 選 **Read and write permissions** → Save

### 步驟三：手動執行看看

`Actions` → `每週千張大戶自動爬取` → `Run workflow` → 按 **Run workflow**

等約 2~3 分鐘，`data/` 資料夾就會出現結果！

---

## ⚙️ 自訂股票清單

編輯 `scripts/scraper.py` 第 28 行的 `DEFAULT_STOCKS`：

```python
DEFAULT_STOCKS = [
    '2330',  # 台積電
    '2317',  # 鴻海
    '2454',  # 聯發科
    # ... 加你想追蹤的
]
```

或手動觸發時在輸入框填入：`2330,2317,2454`

---

## 📅 執行時間

- **自動**：每週五 18:30（台灣時間）
- **手動**：Actions 頁面隨時可觸發

---

## 📊 輸出說明

### Excel 工作表

| 工作表 | 內容 |
|-------|------|
| 持股比%趨勢 | 各股 × 各週千張大戶持股比例 |
| 持股人數趨勢 | 各股 × 各週千張大戶人數 |
| 週增減人數 | 正=大戶加碼 / 負=大戶減碼 |
| 持股比增減% | 週比例變化 |
| 完整原始資料 | 全部15個分級的原始數據 |

### 判讀重點

```
千張大戶比例 ↑ + 總股東人數 ↓ → 籌碼集中，主力默默吃貨 ✅
千張大戶比例 ↓ + 總股東人數 ↑ → 主力出貨給散戶承接 ⚠️
```

---

## 🔑 設定 FINMIND_TOKEN

`total_screener.py`、`screener_to_seeds.py` 以及感測器的 `--auto`／`--full` 都會用 FinMind 抓資料，
資料量大，**務必設 token**（沒設會被限流或退回「待驗證」）。Token 在 [FinMind 會員後台](https://finmindtrade.com/)
產生。**Token 是機密，絕對不要寫進程式碼或 commit 進 repo。**

**本機（當前終端機，臨時）**

```bash
export FINMIND_TOKEN="你的token"      # 注意:此指令本身會留在 shell 歷史
python3 total_screener.py
```

**本機（持久，建議用 .env，已被 .gitignore 忽略）**

```bash
echo 'FINMIND_TOKEN=你的token' >> .env   # .env 不會進 git
set -a; source .env; set +a              # 載入到環境變數後再執行
```

**GitHub Actions（排程/手動觸發用，最安全）**

`Settings` → `Secrets and variables` → `Actions` → `New repository secret`
名稱填 `FINMIND_TOKEN`、值貼上 token。各 workflow 已用 `${{ secrets.FINMIND_TOKEN }}` 讀取，無需改碼。

> ⚠️ 若 token 曾經外流（貼到聊天、訊息、截圖、commit…），請到 FinMind 後台**重新產生**一個，舊的作廢。

---

## 🧭 SLCA 投資感測器 v2（`sensor.py`）

把 `SLCA_______Prompt_v2.md` 的偵測流程程式化：從市場差異產生「種子」，交棒給 SLCA v5 分析。
**感測器只偵測、不分析、不建議買賣**（眼睛不是大腦）。

### 它做什麼

完整實作 Prompt 的七步流程：六種差異掃描 → **差異強度評分 DS**（基礎+共振+歷史−死亡模式）
→ **波普爾三問**假陽性過濾 → 機會成本感測 → **信心分數** → **注意力預算**（A級DS>85最多1顆、
B級DS70–85最多2顆、合計≤3）→ 輸出標準種子格式。內建**死亡模式庫**（敘事泡沫／假底部／
業績轉機幻覺／政策題材），命中即扣分。

### 怎麼跑

```bash
# 1) 產生空白輸入模板
python3 sensor.py --template > sensor_input.json

# 2) 填好當週原始資料（敘事／反共識／法人／波普爾三問答案…），再執行
python3 sensor.py --input sensor_input.json --out data/SLCA_種子.md

# 直接看範例（已附 sensor_input.example.json）
python3 sensor.py --input sensor_input.example.json --out data/SLCA_種子_範例.md

# 選配：額外用 FinMind 自動偵測 ①價格／②基本面／③矛盾（需 FINMIND_TOKEN）
python3 sensor.py --input sensor_input.json --auto

# 走完整 Reality Loop：種子後附「交棒 SLCA v5 指令 + 現實驗證追蹤表 + 演化記錄列」
# --full 會以 total_screener 抓 FinMind 數據回填現實驗證的客觀現況（需 FINMIND_TOKEN，否則退回待驗證）
python3 sensor.py --input sensor_input.json --full --out data/SLCA_種子.md
```

### 從五維總篩選直接出種子（`screener_to_seeds.py`）

已經跑過 `total_screener.py` 的話，可把它的 xlsx **直接橋接成種子**，免再手填輸入：

```bash
pip install openpyxl
python3 screener_to_seeds.py data/台股五維總篩選.xlsx --out data/SLCA_種子_實測.md
```

規則式從五維欄位偵測 ②基本面／③矛盾（便宜的好公司／品質高但股價落後／品質高但燒錢），
純資金假突破套用死亡模式 005；現實驗證的客觀現況直接取自 xlsx（含各檔損益表算毛利率/淨利），
質性欄位（共識／三問）以模板生成並標註「需人工確認」。

> **引擎只用 Python 標準庫**，沒裝 FinMind／pandas 也能跑 —— 量化差異（①②③）可選擇用
> `--auto` 從 FinMind 自動補入，質性差異（④敘事／⑤反共識／⑥時間）與三問答案則由輸入檔提供。
> 也可由 `Actions → SLCA 投資感測器 v2（手動）` 手動觸發。

---

## ⚠️ 注意事項

- 資料來源：[聚財網](https://stock.wearn.com) 集保股權分散表
- 每週六更新，爬蟲設定週五晚間執行（資料為前一週五數據）
- 請勿短時間大量請求，腳本已設 1.5 秒間隔
- 資料僅供參考，不構成投資建議
