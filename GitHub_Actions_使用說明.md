# 🤖 台期所分階段資料爬取 - GitHub Actions 自動化

## 📋 功能說明

本系統使用GitHub Actions實現全自動化的台期所資料爬取，支援兩階段分時爬取：

### 🕐 第一階段：交易量資料 (下午2點)
- **執行時間**: 台灣時間每週一到週五 14:00 (UTC 06:00)
- **資料內容**: 多方交易口數、多方契約金額、空方交易口數、空方契約金額、多空淨額交易口數、多空淨額契約金額
- **Google Sheets分頁**: 「交易量資料」
- **Workflow檔案**: `.github/workflows/crawl_trading_data.yml`

### 🕐 第二階段：完整資料 (下午3點半)
- **執行時間**: 台灣時間每週一到週五 15:30 (UTC 07:30)
- **資料內容**: 交易量資料 + 未平倉資料
- **Google Sheets分頁**: 「完整資料」
- **Workflow檔案**: `.github/workflows/crawl_complete_data.yml`

## 🔧 初始設定

### 1. 設定GitHub Secrets

在GitHub Repository → Settings → Secrets and variables → Actions 中添加以下secrets：

#### 必需的Secrets：
```bash
GOOGLE_SHEETS_CREDENTIALS
# Google服務帳號的JSON憑證（完整的JSON內容）
```

#### 可選的Secrets（如需Telegram通知）：
```bash
TELEGRAM_BOT_TOKEN
# Telegram Bot的Token

TELEGRAM_CHAT_ID  
# Telegram Chat ID
```

### 2. Google Sheets憑證設定

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用 Google Sheets API 和 Google Drive API
4. 建立服務帳號並下載JSON金鑰檔案
5. 將JSON檔案的**完整內容**複製到GitHub Secrets的 `GOOGLE_SHEETS_CREDENTIALS`

### 3. Google Sheets權限設定

將服務帳號的email（在JSON憑證中的`client_email`）添加到您的Google試算表的共享權限中，給予「編輯者」權限。

## 🚀 使用方式

### 自動執行（推薦）

設定完成後，系統會自動在以下時間執行：

- **週一到週五 14:00** (台灣時間): 自動爬取交易量資料
- **週一到週五 15:30** (台灣時間): 自動爬取完整資料

### 手動執行

#### 方法一：GitHub介面操作

1. 前往您的GitHub Repository
2. 點擊 `Actions` 頁籤
3. 選擇要執行的workflow：
   - `🚀 台期所交易量資料爬取`: 爬取交易量資料
   - `📈 台期所完整資料爬取`: 爬取完整資料
   - `⚡ 手動執行台期所資料爬取`: 自訂參數爬取
4. 點擊 `Run workflow` 按鈕

#### 方法二：使用GitHub CLI

```bash
# 爬取交易量資料
gh workflow run "crawl_trading_data.yml"

# 爬取完整資料  
gh workflow run "crawl_complete_data.yml"

# 手動執行並指定參數
gh workflow run "manual_crawl.yml" \
  -f data_type=TRADING \
  -f contracts=TX,TE,MTX \
  -f date_range=today
```

## 📊 執行結果查看

### 1. GitHub Actions頁面
- 前往 Actions 頁籤查看執行狀態
- 點擊特定的workflow run查看詳細日誌
- 下載Artifacts獲取生成的檔案

### 2. Google Sheets
- **交易量資料分頁**: 下午2點的交易量資料
- **完整資料分頁**: 下午3點半的完整資料
- **每日摘要分頁**: 自動生成的統計摘要
- **三大法人趨勢分頁**: 趨勢分析圖表

### 3. Telegram通知（如已設定）
- 自動接收執行成功/失敗通知
- 接收生成的圖表分析

## ⏰ 時間設定說明

### Cron表達式說明
```yaml
# 交易量資料 - 台灣時間14:00
- cron: '0 6 * * 1-5'  # UTC 06:00 = 台灣時間 14:00

# 完整資料 - 台灣時間15:30  
- cron: '30 7 * * 1-5'  # UTC 07:30 = 台灣時間 15:30
```

### 時區轉換
- 台灣時間 (UTC+8) = UTC時間 + 8小時
- GitHub Actions使用UTC時間
- 如需調整時間，請修改對應的workflow檔案中的cron表達式

## 🔍 監控與除錯

### 查看執行日誌
1. 前往 Actions 頁籤
2. 點擊特定的workflow run
3. 展開各個步驟查看詳細日誌

### 常見問題排除

#### 1. Google Sheets認證失敗
```
錯誤: Google Sheets認證失敗
解決: 檢查GOOGLE_SHEETS_CREDENTIALS是否正確設定
```

#### 2. 無法寫入Google Sheets
```
錯誤: Permission denied
解決: 確認服務帳號已被加入試算表的共享權限
```

#### 3. 爬取無資料
```
錯誤: 沒有爬取到任何有效資料
解決: 檢查是否為交易日，以及台期所網站是否正常
```

### 調整執行參數

如需調整爬取參數，可修改workflow檔案中的命令：

```yaml
python taifex_crawler.py \
  --date-range today \
  --contracts TX,TE,MTX,ZMX \    # 添加更多契約
  --identities ALL \
  --data_type TRADING \
  --max_workers 3 \              # 降低並發數
  --delay 2.0                    # 增加延遲
```

## 📈 進階功能

### 自訂執行時間

如需更改執行時間，修改對應workflow檔案中的cron表達式：

```yaml
schedule:
  # 改為台灣時間 14:30 執行
  - cron: '30 6 * * 1-5'  # UTC 06:30 = 台灣時間 14:30
```

### 添加更多契約

修改workflow檔案中的 `--contracts` 參數：

```yaml
--contracts TX,TE,MTX,ZMX,NQF \
```

### 開啟/關閉特定功能

在workflow檔案中調整環境變數：

```yaml
env:
  GOOGLE_SHEETS_CREDENTIALS: ${{ secrets.GOOGLE_SHEETS_CREDENTIALS }}
  # TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}  # 註解掉關閉Telegram通知
  # TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
```

## 🛡️ 安全注意事項

1. **永不在程式碼中存放憑證**: 所有敏感資訊都應存放在GitHub Secrets中
2. **定期更新憑證**: 建議定期更換Google服務帳號金鑰
3. **最小權限原則**: 服務帳號只給予必要的Google Sheets存取權限
4. **監控執行日誌**: 定期檢查是否有異常存取嘗試

## 💡 最佳實踐

1. **測試後部署**: 先使用手動執行測試各項功能
2. **監控執行狀態**: 訂閱GitHub Actions的執行通知
3. **備份重要資料**: 定期備份Google Sheets中的資料
4. **文件化修改**: 記錄任何客製化的修改內容

## 🆘 技術支援

如遇到問題，請：

1. 檢查GitHub Actions的執行日誌
2. 查看Artifacts中的詳細錯誤訊息
3. 確認所有必需的Secrets都已正確設定
4. 參考workflow檔案中的註解說明 