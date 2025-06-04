# Google Sheets 設定指南

## 📋 前置作業

### 1. 建立Google Cloud專案
1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 專案名稱建議：`taiwan-futures-crawler`

### 2. 啟用Google Sheets API
1. 在Cloud Console中，前往「API和服務」>「程式庫」
2. 搜尋「Google Sheets API」
3. 點選並啟用此API

### 3. 建立服務帳號
1. 前往「API和服務」>「憑證」
2. 點選「建立憑證」>「服務帳號」
3. 輸入服務帳號名稱：`taifex-crawler-service`
4. 點選「建立並繼續」
5. 角色選擇：「編輯者」
6. 點選「完成」

### 4. 產生金鑰
1. 找到剛建立的服務帳號
2. 點選「動作」>「管理金鑰」
3. 點選「新增金鑰」>「建立新金鑰」
4. 選擇「JSON」格式
5. 下載金鑰檔案

## 🔐 設定GitHub Secrets

### 1. 複製JSON內容
1. 開啟下載的JSON檔案
2. 複製整個JSON內容

### 2. 設定GitHub Secret
1. 前往你的GitHub倉庫
2. 點選「Settings」
3. 左側選單選擇「Secrets and variables」>「Actions」
4. 點選「New repository secret」
5. 名稱：`GOOGLE_SHEETS_CREDENTIALS`
6. 值：貼上剛複製的JSON內容
7. 點選「Add secret」

## ✅ 驗證設定

重新執行GitHub Actions，應該會看到：
```
✅ Google Sheets認證已設定（來自GitHub Secret）
```

如果看到以下訊息則表示未設定：
```
⚠️ 未設定GOOGLE_SHEETS_CREDENTIALS Secret，跳過Google Sheets功能
```

## 📱 首次使用

設定完成後，程式會自動：
1. 建立新的Google試算表
2. 設定為公開可檢視
3. 儲存試算表連結到 `config/spreadsheet_config.json`
4. 上傳最新資料到試算表

## 🔗 存取試算表

完成後可以在以下位置找到試算表連結：
- GitHub Actions執行日誌中
- `config/spreadsheet_config.json` 檔案中
- 或查看 `execution_summary.txt`

## 🎯 完成後的優勢

✅ **任何時間存取** - 24/7雲端存取
✅ **任何裝置** - 手機、平板、電腦
✅ **即時更新** - 資料自動同步
✅ **內建圖表** - 直接製作分析圖表
✅ **分享功能** - 輕鬆分享給他人
✅ **完全免費** - Google免費額度充足
✅ **永不離線** - 不受電腦開關機影響

## 📱 使用方式

### 手機存取
1. 下載「Google試算表」APP
2. 登入Google帳號
3. 開啟「台期所資料分析」試算表

### 電腦存取
1. 開啟瀏覽器
2. 前往試算表網址
3. 即可查看最新資料

## 🔄 自動化流程

每日09:00 GitHub Actions會自動：
1. 爬取台期所最新資料
2. 上傳到Google Sheets
3. 更新分析圖表
4. 發送完成通知

完全無需人工干預！
