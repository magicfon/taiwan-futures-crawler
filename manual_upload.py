#!/usr/bin/env python3
"""
手動上傳腳本
手動測試Google Sheets資料上傳功能
"""

import json
from pathlib import Path
from google_sheets_manager import GoogleSheetsManager
from database_manager import TaifexDatabaseManager

def main():
    print("🚀 開始手動上傳...")
    
    # 1. 初始化管理器
    sheets_manager = GoogleSheetsManager()
    db_manager = TaifexDatabaseManager()
    
    # 2. 連接到試算表
    config_file = Path("config/spreadsheet_config.json")
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    spreadsheet_id = config.get('spreadsheet_id')
    sheets_manager.connect_spreadsheet(spreadsheet_id)
    
    print(f"✅ 已連接到試算表: {sheets_manager.spreadsheet.title}")
    
    # 3. 獲取資料庫資料
    recent_data = db_manager.get_recent_data(30)
    summary_data = db_manager.get_daily_summary(30)
    
    print(f"📊 準備上傳:")
    print(f"  - 近30天資料: {len(recent_data)} 筆")
    print(f"  - 摘要資料: {len(summary_data)} 筆")
    
    # 4. 上傳主要資料
    if not recent_data.empty:
        print("📤 上傳歷史資料...")
        try:
            success = sheets_manager.upload_data(recent_data)
            if success:
                print("✅ 歷史資料上傳成功")
            else:
                print("❌ 歷史資料上傳失敗")
        except Exception as e:
            print(f"❌ 歷史資料上傳錯誤: {e}")
    
    # 5. 上傳摘要資料
    if not summary_data.empty:
        print("📤 上傳摘要資料...")
        try:
            success = sheets_manager.upload_summary(summary_data)
            if success:
                print("✅ 摘要資料上傳成功")
            else:
                print("❌ 摘要資料上傳失敗")
        except Exception as e:
            print(f"❌ 摘要資料上傳錯誤: {e}")
        
        print("📤 更新趨勢分析...")
        try:
            success = sheets_manager.update_trend_analysis(summary_data)
            if success:
                print("✅ 趨勢分析更新成功")
            else:
                print("❌ 趨勢分析更新失敗")
        except Exception as e:
            print(f"❌ 趨勢分析更新錯誤: {e}")
    
    # 6. 更新系統資訊
    print("📤 更新系統資訊...")
    try:
        success = sheets_manager.update_system_info()
        if success:
            print("✅ 系統資訊更新成功")
        else:
            print("❌ 系統資訊更新失敗")
    except Exception as e:
        print(f"❌ 系統資訊更新錯誤: {e}")
    
    print(f"\n🌐 Google試算表網址: {sheets_manager.get_spreadsheet_url()}")
    print("🎉 手動上傳完成！")

if __name__ == "__main__":
    main() 