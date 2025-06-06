#!/usr/bin/env python3
"""
Google Sheets診斷腳本
檢查Google Sheets的連接狀況和配置
"""

import json
import sys
from pathlib import Path

def main():
    print("🔍 Google Sheets診斷開始...")
    
    # 1. 檢查是否有Google Sheets模組
    try:
        import gspread
        print("✅ gspread模組已安裝")
    except ImportError:
        print("❌ gspread模組未安裝")
        return
    
    # 2. 檢查配置檔案
    config_file = Path("config/spreadsheet_config.json")
    if config_file.exists():
        print("✅ 試算表配置檔案存在")
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print(f"📋 試算表ID: {config.get('spreadsheet_id')}")
        print(f"🔗 網址: {config.get('spreadsheet_url')}")
    else:
        print("❌ 試算表配置檔案不存在")
        return
    
    # 3. 檢查Google憑證
    cred_paths = [
        "config/google_sheets_credentials.json",
        "config/google_credentials.json",
        "config/service_account.json",
        "google_credentials.json"
    ]
    
    cred_found = False
    for cred_path in cred_paths:
        if Path(cred_path).exists():
            print(f"✅ 找到Google憑證: {cred_path}")
            cred_found = True
            break
    
    if not cred_found:
        print("❌ 未找到Google憑證檔案")
        return
    
    # 4. 嘗試連接Google Sheets
    try:
        from google_sheets_manager import GoogleSheetsManager
        sheets_manager = GoogleSheetsManager()
        
        if sheets_manager.client:
            print("✅ Google Sheets認證成功")
            
            # 嘗試連接到指定的試算表
            spreadsheet_id = config.get('spreadsheet_id')
            if spreadsheet_id:
                try:
                    sheets_manager.connect_spreadsheet(spreadsheet_id)
                    if sheets_manager.spreadsheet:
                        print(f"✅ 成功連接到試算表: {sheets_manager.spreadsheet.title}")
                        
                        # 檢查工作表
                        worksheets = sheets_manager.spreadsheet.worksheets()
                        print(f"📊 工作表數量: {len(worksheets)}")
                        for ws in worksheets:
                            print(f"  - {ws.title}")
                        
                        # 嘗試測試上傳
                        print("\n🧪 測試資料上傳...")
                        test_data = [
                            ["測試日期", "測試契約", "測試身份", "測試數值"],
                            ["2025-06-06", "TX", "外資", "100"]
                        ]
                        
                        try:
                            worksheet = sheets_manager.spreadsheet.worksheet("歷史資料")
                            # 只測試讀取，不實際寫入
                            current_data = worksheet.get_all_values()
                            print(f"✅ 可以讀取歷史資料工作表，目前有 {len(current_data)} 行資料")
                        except Exception as e:
                            print(f"❌ 無法存取歷史資料工作表: {e}")
                        
                    else:
                        print("❌ 無法連接到指定的試算表")
                except Exception as e:
                    print(f"❌ 連接試算表失敗: {e}")
            else:
                print("❌ 配置中沒有試算表ID")
        else:
            print("❌ Google Sheets認證失敗")
            
    except Exception as e:
        print(f"❌ Google Sheets模組載入失敗: {e}")
    
    # 5. 檢查資料庫狀況
    try:
        from database_manager import TaifexDatabaseManager
        db_manager = TaifexDatabaseManager()
        
        recent_data = db_manager.get_recent_data(30)
        summary_data = db_manager.get_daily_summary(30)
        
        print(f"\n💾 資料庫狀況:")
        print(f"  - 近30天資料: {len(recent_data)} 筆")
        print(f"  - 摘要資料: {len(summary_data)} 筆")
        
        if len(recent_data) > 0:
            print("✅ 資料庫有資料可上傳")
        else:
            print("⚠️ 資料庫沒有資料")
            
    except Exception as e:
        print(f"❌ 資料庫檢查失敗: {e}")
    
    print("\n🔍 診斷完成")

if __name__ == "__main__":
    main() 