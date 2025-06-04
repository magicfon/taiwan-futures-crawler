# GitHub Actions 自動產生的檔案

這個專案使用 GitHub Actions 自動執行台期所資料爬取，並會產生以下檔案：

## 📁 自動產生的檔案

### `latest_30days_data.xlsx`
- **說明**: 最新30天的台期所資料匯總
- **更新頻率**: 每日自動更新（當有成功爬取資料時）
- **來源**: 由 `output/台期所最新30天資料.xlsx` 複製而來
- **用途**: 可直接下載查看最新的資料分析

### `execution_summary.txt`
- **說明**: GitHub Actions 執行摘要和日誌片段
- **內容**: 包含執行時間、結果狀態和最後20行日誌
- **更新頻率**: 每次 GitHub Actions 執行後更新
- **用途**: 快速了解自動化爬蟲的執行狀況

### `config/spreadsheet_config.json`
- **說明**: Google Sheets 試算表配置（如果啟用）
- **內容**: 試算表ID和URL等資訊
- **用途**: 記錄自動建立的Google試算表位置

## 🚀 如何使用

1. **查看最新資料**: 直接下載 `latest_30days_data.xlsx`
2. **檢查執行狀況**: 查看 `execution_summary.txt`
3. **存取雲端資料**: 透過 `config/spreadsheet_config.json` 中的URL存取Google試算表

## 📋 注意事項

- 這些檔案會被 GitHub Actions 自動更新
- 請勿手動編輯這些檔案
- 如果需要歷史資料，請查看 GitHub Actions 的執行記錄和 Artifacts 