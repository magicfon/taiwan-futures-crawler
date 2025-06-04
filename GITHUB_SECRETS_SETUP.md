# 🔐 GitHub Secrets 設定指南

## 📋 為什麼需要設定GitHub Secrets？

由於安全考量，Google憑證檔案**不能**直接上傳到GitHub。但GitHub Actions需要這些憑證才能自動更新Google Sheets。解決方案是使用GitHub的**Secrets**功能。

## 🚀 完整設定步驟

### 第一步：取得Google憑證
1. 依照 `GOOGLE_SHEETS_完整設定指南.md` 建立Google服務帳號
2. 下載JSON憑證檔案到本地
3. 打開JSON檔案，**複製完整內容**

### 第二步：設定GitHub Secret
1. 前往你的GitHub倉庫
2. 點擊 **Settings** （設定）分頁
3. 在左側選單點擊 **Secrets and variables** → **Actions**
4. 點擊 **New repository secret**
5. 設定以下內容：
   - **Name**: `GOOGLE_SHEETS_CREDENTIALS`
   - **Secret**: 貼上完整的JSON內容
6. 點擊 **Add secret**

### 第三步：設定Telegram Bot（選擇性）
如果你想要接收圖表推送通知：

#### 建立Telegram Bot
1. 在Telegram中搜尋 `@BotFather`
2. 發送 `/newbot` 命令
3. 設定機器人名稱和用戶名
4. 記下Bot Token（格式：`123456789:ABCdefGHI...`）

#### 取得Chat ID
1. 在Telegram中搜尋你的機器人
2. 對機器人發送任何訊息
3. 在瀏覽器中開啟：`https://api.telegram.org/bot<YOUR_BOT_TOKEN>/getUpdates`
4. 找到 `"chat":{"id":12345678}` 部分的數字

#### 設定Telegram Secrets
回到GitHub Secrets頁面，新增：
- **Name**: `TELEGRAM_BOT_TOKEN`
- **Secret**: 你的Bot Token
- **Name**: `TELEGRAM_CHAT_ID`  
- **Secret**: 你的Chat ID

### 第四步：驗證設定
1. 前往 **Actions** 分頁
2. 手動觸發 **每日台期所資料爬取** workflow
3. 檢查執行日誌，應該看到：
   ```
   ✅ Google Sheets認證已設定（來自GitHub Secret）
   ✅ Telegram通知已設定（來自GitHub Secrets）
   ```

## 🔍 範例JSON格式

你的Google憑證JSON應該看起來像這樣：
```json
{
  "type": "service_account",
  "project_id": "your-project-123456",
  "private_key_id": "abcd1234...",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BA...\n-----END PRIVATE KEY-----\n",
  "client_email": "your-service@your-project.iam.gserviceaccount.com",
  "client_id": "123456789012345678901",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/your-service%40your-project.iam.gserviceaccount.com"
}
```

## ✅ 設定完成後的效果

### 🤖 GitHub Actions 自動化
- 每日自動爬取台期所資料
- 自動更新Google Sheets
- 自動發送Telegram圖表報告
- 完全無需人工干預

### ☁️ 雲端存取優勢
- **任何裝置**：手機、平板、電腦都能查看
- **即時更新**：資料自動同步
- **永不離線**：不受個人電腦開關機影響
- **分享便利**：一個連結即可分享給他人

### 📱 Telegram通知優勢
- **即時推送**：圖表生成後立即推送
- **隨身攜帶**：手機隨時查看分析
- **歷史保存**：Telegram自動保存所有圖表
- **分享容易**：直接轉發圖表給他人

## 🔧 故障排除

### 問題1：看到 "未設定GOOGLE_SHEETS_CREDENTIALS Secret"
**解決方案**：
1. 確認Secret名稱完全正確：`GOOGLE_SHEETS_CREDENTIALS`
2. 檢查JSON內容是否完整複製
3. 重新觸發workflow

### 問題2：認證失敗錯誤
**解決方案**：
1. 確認Google Sheets API已啟用
2. 確認Google Drive API已啟用
3. 檢查服務帳號權限
4. 重新下載JSON憑證

### 問題3：無法存取試算表
**解決方案**：
1. 確認試算表已建立
2. 檢查服務帳號email是否有試算表存取權限
3. 或設定試算表為公開可檢視

### 問題4：Telegram無法收到通知
**解決方案**：
1. 確認Bot Token格式正確
2. 確認已對機器人發送過訊息
3. 檢查Chat ID是否正確
4. 確認機器人沒有被封鎖

## 📱 使用指南

### 本地開發
- Google憑證檔案放在 `config/google_sheets_credentials.json`
- Telegram設定可在程式中硬編碼或使用環境變數
- 系統會優先使用本地檔案和設定

### GitHub Actions
- 所有憑證從GitHub Secrets載入
- 自動建立臨時憑證檔案
- 執行完畢後自動清除敏感資料

## 🎯 最佳實踐

1. **定期更新憑證**：Google建議每90天更新一次
2. **權限最小化**：只給服務帳號必要權限
3. **監控使用量**：檢查Google API使用配額
4. **備份重要資料**：定期下載試算表備份
5. **保護Bot Token**：不要在公開場所分享Telegram Bot Token

## 🚀 功能配置總覽

| 功能 | 必需的Secrets | 效果 |
|------|---------------|------|
| 基本爬蟲 | 無 | 爬取資料並儲存到本地檔案 |
| Google Sheets同步 | `GOOGLE_SHEETS_CREDENTIALS` | 雲端自動同步、多裝置存取 |
| Telegram圖表推送 | `TELEGRAM_BOT_TOKEN`<br>`TELEGRAM_CHAT_ID` | 手機即時接收分析圖表 |

**建議配置**：設定所有Secrets，享受完整的自動化體驗！

## 📞 需要協助？

如果設定過程中遇到問題：
1. 檢查GitHub Actions執行日誌
2. 確認所有步驟都正確完成
3. 參考Google Cloud Console錯誤訊息
4. 檢查Telegram機器人狀態

完成設定後，你就有了一個完全自動化的台期所資料分析系統！ 