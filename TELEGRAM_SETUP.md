# Telegram 通知設定指南

## 🤖 建立Telegram Bot

### 1. 找到BotFather
1. 在Telegram中搜尋 `@BotFather`
2. 點選官方的BotFather (有藍色勾勾)
3. 點選「Start」開始對話

### 2. 建立新Bot
1. 發送指令：`/newbot`
2. 輸入Bot名稱：`台期所資料通知機器人`
3. 輸入Bot用戶名：`taifex_data_bot` (必須以_bot結尾)
4. BotFather會回覆Bot Token

### 3. 記錄Bot Token
- 格式類似：`1234567890:ABCdefGHIjklMNOpqrsTUVwxyz`
- **重要**: 請妥善保存，這是機器人的身份認證

## 💬 取得Chat ID

### 方法1: 使用Bot發送訊息
1. 找到你剛建立的Bot
2. 點選「Start」並發送任何訊息
3. 前往瀏覽器開啟：
   ```
   https://api.telegram.org/bot<你的BOT_TOKEN>/getUpdates
   ```
4. 在回應中找到 `"chat":{"id":123456789}` 
5. 記錄這個數字（Chat ID）

### 方法2: 使用@userinfobot
1. 在Telegram中搜尋 `@userinfobot`
2. 點選「Start」
3. 它會顯示你的User ID，這就是Chat ID

## 🔐 設定GitHub Secrets

### 1. 設定Bot Token
1. 前往GitHub倉庫的「Settings」
2. 左側選單選擇「Secrets and variables」>「Actions」
3. 點選「New repository secret」
4. 名稱：`TELEGRAM_BOT_TOKEN`
5. 值：輸入你的Bot Token
6. 點選「Add secret」

### 2. 設定Chat ID
1. 繼續在Secrets頁面
2. 點選「New repository secret」
3. 名稱：`TELEGRAM_CHAT_ID`
4. 值：輸入你的Chat ID
5. 點選「Add secret」

## ✅ 驗證設定

重新執行GitHub Actions，應該會看到：
```
✅ Telegram通知已設定（來自GitHub Secrets）
```

如果看到以下訊息則表示未設定：
```
⚠️ 未設定Telegram Secrets，跳過Telegram通知功能
```

## 📊 功能說明

設定完成後，Telegram Bot會自動發送：

### 每日通知內容
- 📈 **資料摘要圖表**
- 📊 **趨勢分析圖**
- 💹 **各契約成交量對比**
- 🏛️ **三大法人部位變化**

### 圖表類型
1. **成交量趨勢圖** - 30天歷史成交量變化
2. **法人部位圖** - 自營商、投信、外資部位對比
3. **契約分析圖** - TX、TE、MTX等各契約表現
4. **淨部位變化圖** - 多空部位變化趨勢

## 🔧 測試功能

如果要測試Telegram功能：
1. 手動觸發GitHub Actions
2. 使用參數設定最近幾天的日期範圍
3. 確認有足夠的歷史資料生成圖表

## 📱 接收通知

設定完成後：
- 每天台期所資料更新時會自動推送
- 包含圖表和文字摘要
- 可以直接在Telegram中查看分析結果

## 🚫 停用通知

如果不想接收通知：
1. 刪除GitHub Secrets中的Telegram設定
2. 或者在Telegram中封鎖Bot
3. 或者發送 `/stop` 給Bot 