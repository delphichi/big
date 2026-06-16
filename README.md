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

## ⚠️ 注意事項

- 資料來源：[聚財網](https://stock.wearn.com) 集保股權分散表
- 每週六更新，爬蟲設定週五晚間執行（資料為前一週五數據）
- 請勿短時間大量請求，腳本已設 1.5 秒間隔
- 資料僅供參考，不構成投資建議
