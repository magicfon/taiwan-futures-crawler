#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Sheets 資料管理系統
自動將台期所資料上傳到Google試算表
完全雲端化，任何裝置都能存取
"""

import pandas as pd
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
import logging
import time

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    print("警告: Google Sheets套件未安裝，請執行: pip install gspread google-auth")

class GoogleSheetsManager:
    """Google Sheets 管理器"""
    
    def __init__(self, credentials_file="config/google_sheets_credentials.json"):
        """
        初始化Google Sheets管理器
        
        Args:
            credentials_file: Google服務帳號認證檔案路徑
        """
        self.credentials_file = Path(credentials_file)
        self.logger = logging.getLogger(__name__)
        self.client = None
        self.spreadsheet = None
        self.spreadsheet_url = None
        
        if GSPREAD_AVAILABLE:
            self.setup_credentials()
    
    def setup_credentials(self):
        """設定Google Sheets認證"""
        # 優先從環境變數讀取憑證（適用於GitHub Actions）
        credentials_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        
        if credentials_json:
            try:
                self.logger.info("從環境變數載入Google憑證")
                credentials_data = json.loads(credentials_json)
                
                # 設定權限範圍
                scopes = [
                    'https://www.googleapis.com/auth/spreadsheets',
                    'https://www.googleapis.com/auth/drive'
                ]
                
                # 從字典建立認證
                credentials = Credentials.from_service_account_info(
                    credentials_data, 
                    scopes=scopes
                )
                
                # 建立gspread客戶端
                self.client = gspread.authorize(credentials)
                self.logger.info("Google Sheets認證成功（來自環境變數）")
                return True
                
            except Exception as e:
                self.logger.error(f"從環境變數載入Google憑證失敗: {e}")
                # 繼續嘗試從檔案載入
        
        # 從檔案載入憑證（本地開發）
        if not self.credentials_file.exists():
            self.create_credentials_template()
            return False
        
        try:
            # 設定權限範圍
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # 載入認證資訊
            credentials = Credentials.from_service_account_file(
                self.credentials_file, 
                scopes=scopes
            )
            
            # 建立gspread客戶端
            self.client = gspread.authorize(credentials)
            self.logger.info("Google Sheets認證成功（來自檔案）")
            return True
            
        except Exception as e:
            self.logger.error(f"Google Sheets認證失敗: {e}")
            return False
    
    def create_credentials_template(self):
        """建立認證檔案模板"""
        template = {
            "type": "service_account",
            "project_id": "your-project-id",
            "private_key_id": "your-private-key-id",
            "private_key": "-----BEGIN PRIVATE KEY-----\nyour-private-key\n-----END PRIVATE KEY-----\n",
            "client_email": "your-service-account@your-project.iam.gserviceaccount.com",
            "client_id": "your-client-id",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project.iam.gserviceaccount.com"
        }
        
        self.credentials_file.parent.mkdir(exist_ok=True)
        with open(self.credentials_file, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2)
        
        print(f"📋 請完成Google Sheets設定:")
        print(f"1. 前往 https://console.cloud.google.com/")
        print(f"2. 建立專案並啟用Google Sheets API")
        print(f"3. 建立服務帳號並下載JSON金鑰")
        print(f"4. 將內容貼到: {self.credentials_file}")
        print(f"5. 重新執行此程式")
    
    def create_spreadsheet(self, title="台期所資料分析"):
        """建立新的Google試算表"""
        if not self.client:
            self.logger.error("Google Sheets未連接")
            return None
        
        try:
            # 建立新試算表
            spreadsheet = self.client.create(title)
            self.spreadsheet = spreadsheet
            self.spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
            
            # 設定工作表
            self.setup_worksheets()
            
            self.logger.info(f"試算表建立成功: {self.spreadsheet_url}")
            return spreadsheet
            
        except Exception as e:
            self.logger.error(f"建立試算表失敗: {e}")
            return None
    
    def connect_spreadsheet(self, spreadsheet_id_or_url):
        """連接到現有的Google試算表"""
        if not self.client:
            self.logger.error("Google Sheets未連接")
            return None
        
        try:
            # 從URL中提取ID
            if 'docs.google.com' in spreadsheet_id_or_url:
                spreadsheet_id = spreadsheet_id_or_url.split('/d/')[1].split('/')[0]
            else:
                spreadsheet_id = spreadsheet_id_or_url
            
            # 連接試算表
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            self.spreadsheet = spreadsheet
            self.spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
            
            self.logger.info(f"成功連接試算表: {spreadsheet.title}")
            return spreadsheet
            
        except Exception as e:
            self.logger.error(f"連接試算表失敗: {e}")
            return None
    
    def setup_worksheets(self):
        """設定工作表結構"""
        if not self.spreadsheet:
            return
        
        try:
            # 建立主要工作表
            worksheets_config = [
                {"name": "歷史資料", "headers": self.get_history_headers()},
                {"name": "每日摘要", "headers": self.get_summary_headers()},
                {"name": "三大法人趨勢", "headers": self.get_trend_headers()},
                {"name": "系統資訊", "headers": ["項目", "數值", "更新時間"]}
            ]
            
            # 刪除預設工作表
            try:
                default_sheet = self.spreadsheet.worksheet("工作表1")
                self.spreadsheet.del_worksheet(default_sheet)
            except:
                pass
            
            # 建立新工作表
            for config in worksheets_config:
                try:
                    worksheet = self.spreadsheet.add_worksheet(
                        title=config["name"], 
                        rows=1000, 
                        cols=20
                    )
                    
                    # 設定標題行
                    if config["headers"]:
                        worksheet.update('A1', [config["headers"]])
                        
                        # 格式化標題行
                        worksheet.format('A1:Z1', {
                            'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.9},
                            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
                        })
                    
                    self.logger.info(f"工作表建立成功: {config['name']}")
                    
                except Exception as e:
                    self.logger.warning(f"建立工作表 {config['name']} 失敗: {e}")
            
        except Exception as e:
            self.logger.error(f"設定工作表失敗: {e}")
    
    def get_history_headers(self):
        """取得歷史資料表標題"""
        return [
            "日期", "契約名稱", "身份別", 
            "多方交易口數", "多方契約金額", "空方交易口數", "空方契約金額",
            "多空淨額交易口數", "多空淨額契約金額",
            "多方未平倉口數", "多方未平倉契約金額", "空方未平倉口數", "空方未平倉契約金額",
            "多空淨額未平倉口數", "多空淨額未平倉契約金額", "更新時間"
        ]
    
    def get_summary_headers(self):
        """取得摘要表標題"""
        return [
            "日期", "總契約數", "總成交量", 
            "外資淨部位", "自營商淨部位", "投信淨部位", "更新時間"
        ]
    
    def get_trend_headers(self):
        """取得趨勢表標題"""
        return [
            "日期", "外資淨部位", "自營商淨部位", "投信淨部位",
            "外資7日平均", "自營商7日平均", "投信7日平均"
        ]
    
    def upload_data(self, df, worksheet_name="歷史資料"):
        """上傳資料到Google Sheets - 保守的資料管理，不清除歷史資料"""
        if not self.spreadsheet or df.empty:
            return False
        
        try:
            worksheet = self.spreadsheet.worksheet(worksheet_name)
            
            # 檢查現有資料行數
            existing_data = worksheet.get_all_values()
            current_rows = len(existing_data)
            
            # 如果接近Google Sheets的10,000行限制，給出警告但不自動清理
            if current_rows > 9000:
                self.logger.warning(f"⚠️ Google Sheets行數接近限制 ({current_rows} 行)")
                self.logger.warning("建議手動整理資料或建立新的工作表")
                # 不自動清理，讓用戶決定如何處理
            
            # 找到最後一行位置並追加新資料
            last_row = current_rows
            
            # 準備新資料（適應原始爬蟲資料格式）
            data_to_upload = []
            for _, row in df.iterrows():
                data_row = [
                    row.get('日期', ''),
                    row.get('契約名稱', ''),
                    row.get('身份別', ''),
                    row.get('多方交易口數', 0),
                    row.get('多方契約金額', 0),
                    row.get('空方交易口數', 0),
                    row.get('空方契約金額', 0),
                    row.get('多空淨額交易口數', 0),
                    row.get('多空淨額契約金額', 0),
                    row.get('多方未平倉口數', 0),
                    row.get('多方未平倉契約金額', 0),
                    row.get('空方未平倉口數', 0),
                    row.get('空方未平倉契約金額', 0),
                    row.get('多空淨額未平倉口數', 0),
                    row.get('多空淨額未平倉契約金額', 0),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
                data_to_upload.append(data_row)
            
            # 追加新資料到最後一行之後
            if data_to_upload:
                try:
                    start_cell = f'A{last_row + 1}'
                    worksheet.update(start_cell, data_to_upload)
                    self.logger.info(f"成功追加 {len(data_to_upload)} 筆資料到 {worksheet_name}")
                except Exception as upload_error:
                    if "exceeds grid limits" in str(upload_error):
                        self.logger.error("❌ Google Sheets已達行數限制，無法新增更多資料")
                        self.logger.info("💡 建議：手動清理舊資料或建立新的工作表")
                        return False
                    else:
                        raise upload_error
            
            return True
            
        except Exception as e:
            self.logger.error(f"上傳資料到 {worksheet_name} 失敗: {e}")
            return False
    
    def get_recent_data_for_report(self, days=30):
        """從Google Sheets取得最近N天的資料用於派報"""
        if not self.spreadsheet:
            return pd.DataFrame()
        
        try:
            worksheet = self.spreadsheet.worksheet("歷史資料")
            all_data = worksheet.get_all_records()
            
            if not all_data:
                return pd.DataFrame()
            
            # 轉換為DataFrame
            df = pd.DataFrame(all_data)
            
            # 調試輸出
            self.logger.info(f"總共取得 {len(df)} 筆歷史資料")
            if not df.empty:
                self.logger.info(f"日期欄位樣本: {df['日期'].head().tolist()}")
            
            # 轉換日期格式並排序
            try:
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                # 移除無效日期
                df = df.dropna(subset=['日期'])
                df = df.sort_values('日期', ascending=False)
                
                # 取得最近N天的資料
                cutoff_date = datetime.now() - timedelta(days=days)
                recent_df = df[df['日期'] >= cutoff_date]
                
                self.logger.info(f"日期轉換成功，取得最近 {days} 天的資料: {len(recent_df)} 筆")
                return recent_df
                
            except Exception as date_error:
                self.logger.error(f"日期處理失敗: {date_error}")
                # 如果日期處理失敗，回傳原始資料
                self.logger.info(f"回傳原始資料: {len(df)} 筆")
                return df
            
        except Exception as e:
            self.logger.error(f"取得最近資料失敗: {e}")
            return pd.DataFrame()
    
    def upload_summary(self, summary_df):
        """上傳每日摘要資料"""
        if not self.spreadsheet or summary_df.empty:
            return False
        
        try:
            worksheet = self.spreadsheet.worksheet("每日摘要")
            
            # 清除舊資料
            worksheet.batch_clear(["A2:Z1000"])
            
            # 準備摘要資料
            data_to_upload = []
            for _, row in summary_df.iterrows():
                # 確保所有資料都轉換為字串或數字，避免 bytes 類型
                data_row = [
                    str(row.get('date', '')),
                    int(row.get('total_contracts', 0)) if pd.notna(row.get('total_contracts', 0)) else 0,
                    int(row.get('total_volume', 0)) if pd.notna(row.get('total_volume', 0)) else 0,
                    int(row.get('foreign_net', 0)) if pd.notna(row.get('foreign_net', 0)) else 0,
                    int(row.get('dealer_net', 0)) if pd.notna(row.get('dealer_net', 0)) else 0,
                    int(row.get('trust_net', 0)) if pd.notna(row.get('trust_net', 0)) else 0,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
                data_to_upload.append(data_row)
            
            # 上傳資料
            if data_to_upload:
                worksheet.update('A2', data_to_upload)
                self.logger.info(f"成功上傳 {len(data_to_upload)} 筆摘要資料")
            
            return True
            
        except Exception as e:
            self.logger.error(f"上傳摘要資料失敗: {e}")
            return False
    
    def update_trend_analysis(self, summary_df):
        """更新三大法人趨勢分析"""
        if not self.spreadsheet or summary_df.empty:
            return False
        
        try:
            worksheet = self.spreadsheet.worksheet("三大法人趨勢")
            
            # 檢查並轉換數值欄位
            numeric_columns = ['foreign_net', 'dealer_net', 'trust_net']
            for col in numeric_columns:
                if col in summary_df.columns:
                    summary_df[col] = pd.to_numeric(summary_df[col], errors='coerce').fillna(0)
            
            # 排序並檢查是否有足夠資料
            summary_df = summary_df.sort_values('date')
            
            if len(summary_df) == 0:
                self.logger.warning("沒有有效的摘要資料用於趨勢分析")
                return False
            
            # 計算移動平均（至少需要1筆資料）
            if len(summary_df) >= 1:
                summary_df['外資7日平均'] = summary_df['foreign_net'].rolling(7, min_periods=1).mean().round(2)
                summary_df['自營商7日平均'] = summary_df['dealer_net'].rolling(7, min_periods=1).mean().round(2)
                summary_df['投信7日平均'] = summary_df['trust_net'].rolling(7, min_periods=1).mean().round(2)
            else:
                # 如果沒有足夠資料，設為0
                summary_df['外資7日平均'] = 0
                summary_df['自營商7日平均'] = 0  
                summary_df['投信7日平均'] = 0
            
            # 清除舊資料
            worksheet.batch_clear(["A2:Z1000"])
            
            # 準備趨勢資料
            data_to_upload = []
            for _, row in summary_df.iterrows():
                data_row = [
                    str(row.get('date', '')),
                    int(row.get('foreign_net', 0)) if pd.notna(row.get('foreign_net', 0)) else 0,
                    int(row.get('dealer_net', 0)) if pd.notna(row.get('dealer_net', 0)) else 0,
                    int(row.get('trust_net', 0)) if pd.notna(row.get('trust_net', 0)) else 0,
                    round(float(row.get('外資7日平均', 0)), 2) if pd.notna(row.get('外資7日平均', 0)) else 0,
                    round(float(row.get('自營商7日平均', 0)), 2) if pd.notna(row.get('自營商7日平均', 0)) else 0,
                    round(float(row.get('投信7日平均', 0)), 2) if pd.notna(row.get('投信7日平均', 0)) else 0
                ]
                data_to_upload.append(data_row)
            
            # 上傳資料
            if data_to_upload:
                worksheet.update('A2', data_to_upload)
                self.logger.info(f"成功更新趨勢分析資料")
            
            return True
            
        except Exception as e:
            self.logger.error(f"更新趨勢分析失敗: {e}")
            return False
    
    def update_system_info(self):
        """更新系統資訊"""
        if not self.spreadsheet:
            return False
        
        try:
            worksheet = self.spreadsheet.worksheet("系統資訊")
            
            # 系統資訊
            info_data = [
                ["最後更新時間", datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "自動更新"],
                ["資料來源", "台灣期貨交易所", "官方網站"],
                ["更新頻率", "每日09:00", "GitHub Actions"],
                ["試算表網址", self.spreadsheet_url, "可分享連結"],
                ["系統狀態", "運行中", "正常"],
                ["", "", ""],
                ["使用說明", "", ""],
                ["手機存取", "開啟Google試算表APP", ""],
                ["電腦存取", "直接開啟網頁連結", ""],
                ["資料分析", "可直接在試算表中製作圖表", ""],
                ["分享功能", "點擊右上角分享按鈕", ""]
            ]
            
            # 清除並更新
            worksheet.clear()
            worksheet.update('A1', [["項目", "數值", "說明"]])
            worksheet.update('A2', info_data)
            
            # 格式化
            worksheet.format('A1:C1', {
                'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.9},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
            })
            
            return True
            
        except Exception as e:
            self.logger.error(f"更新系統資訊失敗: {e}")
            return False
    
    def create_charts(self):
        """在Google Sheets中建立圖表"""
        if not self.spreadsheet:
            return False
        
        try:
            # 這裡可以使用Google Sheets API建立圖表
            # 由於圖表API較複雜，建議手動在Google Sheets中建立
            self.logger.info("請手動在Google Sheets中建立圖表")
            return True
            
        except Exception as e:
            self.logger.error(f"建立圖表失敗: {e}")
            return False
    
    def get_spreadsheet_url(self):
        """取得試算表網址"""
        return self.spreadsheet_url
    
    def share_spreadsheet(self, email=None, role='reader'):
        """分享試算表"""
        if not self.spreadsheet:
            return False
        
        try:
            if email:
                self.spreadsheet.share(email, perm_type='user', role=role)
                self.logger.info(f"試算表已分享給: {email}")
            else:
                # 設定為公開可檢視
                self.spreadsheet.share('', perm_type='anyone', role='reader')
                self.logger.info("試算表已設定為公開可檢視")
            
            return True
            
        except Exception as e:
            self.logger.error(f"分享試算表失敗: {e}")
            return False


def create_setup_guide():
    """建立Google Sheets設定指南"""
    guide_content = """
# 🌐 Google Sheets 設定指南

## 📋 快速設定步驟

### 1. 建立Google Cloud專案
1. 前往 https://console.cloud.google.com/
2. 點擊「建立專案」
3. 輸入專案名稱：「台期所資料分析」
4. 點擊「建立」

### 2. 啟用Google Sheets API
1. 在專案中搜尋「Google Sheets API」
2. 點擊「啟用」
3. 同樣啟用「Google Drive API」

### 3. 建立服務帳號
1. 前往「IAM與管理」→「服務帳號」
2. 點擊「建立服務帳號」
3. 輸入名稱：「sheets-automation」
4. 點擊「建立並繼續」
5. 略過權限設定，直接點擊「完成」

### 4. 下載認證金鑰
1. 點擊剛建立的服務帳號
2. 切換到「金鑰」分頁
3. 點擊「新增金鑰」→「建立新金鑰」
4. 選擇「JSON」格式
5. 下載檔案並重新命名為：`google_sheets_credentials.json`
6. 將檔案放到 `config/` 目錄下

### 5. 測試連接
```bash
python google_sheets_manager.py
```

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
"""
    
    Path("GOOGLE_SHEETS_SETUP.md").write_text(guide_content, encoding='utf-8')
    print("📖 設定指南已建立: GOOGLE_SHEETS_SETUP.md")


if __name__ == "__main__":
    if not GSPREAD_AVAILABLE:
        print("請先安裝必要套件:")
        print("pip install gspread google-auth")
        create_setup_guide()
        exit(1)
    
    # 測試Google Sheets連接
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets認證失敗")
        print("請完成設定後重新執行")
        create_setup_guide()
    else:
        print("✅ Google Sheets連接成功！")
        print("現在可以建立試算表並上傳資料了") 