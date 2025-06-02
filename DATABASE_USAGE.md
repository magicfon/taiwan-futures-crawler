# 🗄️ 台期所資料庫系統使用指南

## 📋 系統概述

新的資料管理系統提供以下功能：

1. **SQLite本地資料庫** - 所有資料累積儲存，支援快速查詢和分析
2. **固定30天檔案** - 自動更新的最新30天資料Excel檔案
3. **智能日報系統** - 自動生成分析報告和圖表
4. **雲端備份支援** - 支援Supabase等雲端資料庫（可選）

## 📂 新的檔案結構

```
📦 專案根目錄
├── 📁 data/                          # 資料庫目錄
│   └── 📄 taifex_data.db             # SQLite主資料庫
├── 📁 output/                        # 輸出目錄
│   ├── 📄 台期所最新30天資料.xlsx      # 固定更新的30天資料
│   └── 📄 其他歷史資料檔案...
├── 📁 reports/                       # 分析報告目錄
│   ├── 📄 台期所30日報告_20241225.json
│   ├── 📄 台期所30日報告_20241225.xlsx
│   ├── 📄 台期所30日報告_20241225.md
│   └── 📁 charts_20241225/           # 圖表目錄
├── 📁 backup/                        # 備份目錄
│   └── 📄 futures_data_20241225_120000.csv
└── 📁 config/                        # 設定目錄
    └── 📄 cloud_db.json              # 雲端資料庫設定
```

## 🎯 核心功能說明

### 1. 自動資料累積

- ✅ **統一儲存格式**：所有歷史資料存入SQLite資料庫
- ✅ **自動去重**：避免重複資料插入
- ✅ **快速查詢**：支援按日期、契約、身份別快速查詢
- ✅ **資料完整性**：自動驗證和修復資料

### 2. 固定30天檔案

**檔案名稱**：`台期所最新30天資料.xlsx`

**更新頻率**：每次爬蟲執行後自動更新

**內容包含**：
- 📊 **原始資料** - 最近30天完整交易資料
- 📈 **每日摘要** - 三大法人每日淨部位統計
- 📉 **三大法人透視表** - 按日期和身份別交叉分析

### 3. 智能日報系統

**自動生成內容**：
- 📋 基本資訊統計
- 🏛️ 三大法人分析
- 📊 契約活躍度分析
- 📈 趨勢判斷
- ⚠️ 風險警示

**輸出格式**：
- JSON（程式化處理）
- Excel（資料分析）
- Markdown（可讀報告）
- PNG圖表（視覺化）

## 🚀 使用方式

### 基本爬蟲（傳統方式）
```bash
python taifex_crawler.py --date-range today --contracts TX,TE,MTX
```

### 資料庫查詢
```python
from database_manager import TaifexDatabaseManager

# 初始化資料庫
db = TaifexDatabaseManager()

# 查詢最近30天資料
recent_data = db.get_recent_data(30)

# 查詢每日摘要
summary = db.get_daily_summary(30)

# 匯出Excel
db.export_to_excel("my_analysis.xlsx", days=30)
```

### 生成分析報告
```python
from daily_report_generator import DailyReportGenerator

# 生成30天報告
generator = DailyReportGenerator()
report = generator.generate_30day_report()
```

## 📊 推薦的免費資料庫方案

### 方案一：純本地方案（推薦新手）
- **資料庫**：SQLite（本地檔案）
- **備份**：GitHub自動版本控制
- **優點**：設定簡單，完全免費，無網路限制
- **缺點**：無法多設備同步

### 方案二：本地+雲端混合（推薦進階用戶）
- **主資料庫**：SQLite（本地）
- **雲端備份**：Supabase免費層
- **每日同步**：自動上傳關鍵資料
- **優點**：本地快速，雲端備份，支援多設備
- **免費額度**：500MB資料庫，5萬次API請求/月

### 方案三：完全雲端方案
- **資料庫**：MongoDB Atlas免費層
- **備份**：自動雲端備份
- **優點**：多設備同步，無需本地維護
- **免費額度**：512MB儲存，共享叢集

## ⚙️ 雲端資料庫設定（Supabase範例）

### 1. 註冊Supabase帳號
前往 [supabase.com](https://supabase.com) 註冊免費帳號

### 2. 建立專案
- 建立新專案
- 記下專案URL和API金鑰

### 3. 設定本地配置
修改 `config/cloud_db.json`：
```json
{
  "supabase": {
    "url": "https://your-project.supabase.co",
    "key": "your-anon-key",
    "enabled": true
  },
  "backup_enabled": true,
  "sync_interval_hours": 24
}
```

### 4. 建立資料表
在Supabase SQL編輯器中執行：
```sql
CREATE TABLE futures_data (
  id SERIAL PRIMARY KEY,
  date TEXT NOT NULL,
  contract_code TEXT NOT NULL,
  identity_type TEXT NOT NULL,
  position_type TEXT NOT NULL,
  long_position INTEGER,
  short_position INTEGER,
  net_position INTEGER,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

CREATE UNIQUE INDEX idx_unique_record 
ON futures_data(date, contract_code, identity_type, position_type);
```

## 🔄 資料分析工作流程

### 每日自動化流程
1. **09:00** - GitHub Actions自動觸發爬蟲
2. **09:05** - 資料存入SQLite資料庫
3. **09:06** - 更新固定30天檔案
4. **09:07** - 生成分析報告（如果資料足夠）
5. **09:08** - 提交所有變更到GitHub

### 手動分析流程
1. **下載最新資料**：直接查看 `台期所最新30天資料.xlsx`
2. **深度分析**：使用reports目錄中的完整報告
3. **自定義查詢**：使用Python直接查詢資料庫
4. **歷史對比**：查詢資料庫中的長期歷史資料

## 📈 資料分析範例

### 查詢外資近期趨勢
```python
# 取得外資最近30天淨部位
foreign_data = db.get_recent_data(30)
foreign_net = foreign_data[foreign_data['identity_type'] == '外資']['net_position']

# 計算趨勢
trend = "上升" if foreign_net.tail(7).mean() > foreign_net.head(7).mean() else "下降"
print(f"外資近期趨勢：{trend}")
```

### 分析契約活躍度
```python
# 各契約成交量排行
contract_volume = recent_data.groupby('contract_code')['long_position'].sum().sort_values(ascending=False)
print("契約活躍度排行：")
print(contract_volume.head())
```

## ⚠️ 注意事項

1. **資料庫檔案備份**：定期備份 `data/taifex_data.db`
2. **磁碟空間管理**：定期清理舊的分析報告
3. **雲端同步頻率**：避免過於頻繁的API呼叫
4. **資料驗證**：定期檢查資料完整性

## 🆘 常見問題

### Q: 資料庫檔案太大怎麼辦？
A: SQLite支援TB級別資料，正常使用下幾年內不會有問題。如需要可以定期匯出舊資料並清理。

### Q: 如何恢復資料庫？
A: 從 `backup/` 目錄找到最近的CSV備份檔案，重新匯入資料庫。

### Q: 雲端同步失敗怎麼辦？
A: 檢查網路連接和API配置，同步失敗不影響本地功能。

### Q: 如何自定義分析指標？
A: 修改 `daily_report_generator.py` 中的分析邏輯，添加您需要的指標。

---

**系統優勢總結**：
- ✅ 資料永久累積，支援長期分析
- ✅ 固定檔案格式，便於第三方工具使用
- ✅ 自動化分析，減少人工處理
- ✅ 多層備份，確保資料安全
- ✅ 完全免費，無月費負擔 