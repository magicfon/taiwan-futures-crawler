# 📊 遺漏資料檢查和補齊系統

## 🎯 系統目標

在每日GitHub Actions爬取執行之前，自動檢查近30天內是否有遺漏的交易日資料，並自動補齊到資料庫和Google Sheets。

## 🔧 核心組件

### 1. 遺漏資料檢查腳本 (`check_and_fill_missing_data.py`)

**功能：**
- 檢查近N天內應有的交易日（週一到週五）
- 比對資料庫中現有的日期資料
- 自動爬取遺漏的交易日資料
- 同步資料到Google Sheets
- 生成詳細的檢查報告

**使用方式：**
```bash
# 檢查近30天並同步到Google Sheets
python check_and_fill_missing_data.py --days 30

# 只檢查和補齊，跳過Google Sheets同步
python check_and_fill_missing_data.py --days 30 --skip-sync

# 檢查近7天
python check_and_fill_missing_data.py --days 7
```

### 2. GitHub Actions 工作流程 (`.github/workflows/daily_crawler.yml`)

**更新內容：**
- 添加了遺漏資料檢查步驟（在爬蟲執行之前）
- 新增了30天資料圖表和報告生成
- 增強了檔案提交和推送邏輯
- 改進了日誌和報告收集

**執行順序：**
1. 🔍 **檢查並補齊遺漏資料** (新增)
2. 🚀 **執行今日爬蟲**
3. 📊 **生成30天資料圖表和報告** (新增)
4. 💾 **提交和推送更新**

### 3. 測試工作流程腳本 (`test_missing_check_workflow.py`)

**功能：**
- 模擬完整的GitHub Actions執行流程
- 本地測試遺漏資料檢查功能
- 驗證所有組件的整合性

## 📋 執行流程詳解

### 步驟 1: 遺漏資料檢查
```
🔍 檢查近30天遺漏的交易日資料...
├── 計算期間內的交易日 (排除週末)
├── 查詢資料庫現有日期
├── 找出遺漏的交易日
└── 生成遺漏清單
```

### 步驟 2: 自動補齊資料
```
🚀 開始補齊遺漏資料...
├── 逐日執行爬蟲 (TX, TE, MTX)
├── 儲存到本地資料庫
├── 處理爬取失敗的情況
└── 記錄補齊結果
```

### 步驟 3: 同步到Google Sheets
```
📊 同步資料到Google Sheets...
├── 讀取近30天資料
├── 轉換為Sheets格式
├── 上傳到歷史資料工作表
└── 更新時間戳記
```

### 步驟 4: 生成報告
```
📝 生成檢查報告...
├── 記錄檢查時間和範圍
├── 統計遺漏和補齊情況
├── 評估資料庫健康狀態
└── 儲存為JSON報告
```

## 📊 報告格式

**遺漏資料檢查報告** (`reports/missing_data_check_report.json`)：
```json
{
  "check_time": "2025-06-07 13:45:29",
  "check_period_days": 30,
  "missing_dates_count": 3,
  "missing_dates": [
    "2025-06-02",
    "2025-06-03", 
    "2025-06-04"
  ],
  "database_status": "has_gaps",
  "google_sheets_status": "enabled"
}
```

## 🔧 配置選項

### GitHub Actions 輸入參數

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `skip_missing_check` | 跳過遺漏資料檢查 | `false` |
| `date_range` | 今日爬取的日期範圍 | `today` |
| `contracts` | 契約代碼清單 | `TX,TE,MTX,ZMX,NQF` |
| `retry_if_empty` | 沒有資料時是否重試 | `true` |

### 環境要求

**必要密鑰：**
- `GOOGLE_SHEETS_CREDENTIALS`: Google Sheets API 認證
- `TELEGRAM_BOT_TOKEN`: Telegram 機器人權杖 (選用)
- `TELEGRAM_CHAT_ID`: Telegram 聊天室ID (選用)

## 🚀 部署和使用

### 1. 本地測試
```bash
# 測試遺漏資料檢查
python check_and_fill_missing_data.py --days 10 --skip-sync

# 測試完整工作流程
python test_missing_check_workflow.py
```

### 2. GitHub Actions 手動觸發
1. 進入 GitHub Repository
2. 點選 `Actions` 頁籤
3. 選擇 `每日台期所資料爬取`
4. 點選 `Run workflow`
5. 設定參數並執行

### 3. 自動執行時間表
- **每日台北時間 15:30** (UTC 07:30)
- 交易時間結束後，確保資料已發布

## 📈 監控和維護

### 檢查點
1. **每日執行日誌** - GitHub Actions 執行記錄
2. **遺漏資料報告** - 定期檢查 `reports/` 目錄
3. **資料庫完整性** - 使用 `check_db_data.py` 驗證
4. **Google Sheets同步** - 確認 `歷史資料` 工作表更新

### 故障排除

**常見問題：**

1. **遺漏資料檢查失敗**
   - 檢查網路連線
   - 確認台期所網站可訪問
   - 查看 `missing_data_check.log`

2. **Google Sheets 同步失敗**  
   - 驗證 `GOOGLE_SHEETS_CREDENTIALS` 密鑰
   - 檢查工作表權限設定
   - 確認工作表名稱正確

3. **資料格式問題**
   - 檢查日期格式一致性
   - 確認資料庫欄位類型
   - 查看資料轉換日誌

## 🔮 未來改進

- [ ] 支援假日和特殊交易日排除清單
- [ ] 智慧重試機制（根據失敗原因）
- [ ] 增加Email通知功能
- [ ] 資料品質驗證（數值合理性檢查）
- [ ] 支援更多契約類型的自動檢測
- [ ] 增加資料備份和恢復功能

## 📞 支援

如有問題或建議，請檢查：
1. 執行日誌檔案
2. GitHub Actions 執行記錄  
3. 資料庫和Google Sheets狀態
4. 系統時間和時區設定 