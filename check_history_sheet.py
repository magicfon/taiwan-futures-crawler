#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
檢查Google Sheets歷史資料工作表內容
"""

from google_sheets_manager import GoogleSheetsManager
import json
from pathlib import Path

def main():
    print("🔍 檢查Google Sheets歷史資料工作表...")
    
    # 載入配置
    config_file = Path('config/spreadsheet_config.json')
    if not config_file.exists():
        print("❌ 找不到試算表配置檔案")
        return
    
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # 連接Google Sheets
    manager = GoogleSheetsManager()
    if not manager.client:
        print("❌ Google Sheets認證失敗")
        return
    
    manager.connect_spreadsheet(config['spreadsheet_id'])
    if not manager.spreadsheet:
        print("❌ 無法連接到試算表")
        return
    
    print(f"✅ 已連接到試算表: {manager.spreadsheet.title}")
    
    # 檢查歷史資料工作表
    try:
        worksheet = manager.spreadsheet.worksheet('歷史資料')
        data = worksheet.get_all_values()
        
        print(f"\n📊 歷史資料工作表狀況:")
        print(f"   總行數: {len(data)} 行")
        
        if data:
            print(f"   標題行: {data[0]}")
            
            # 尋找最近的資料
            print(f"\n📅 最近10行資料:")
            recent_data = data[-10:] if len(data) > 10 else data[1:]
            
            for i, row in enumerate(recent_data):
                if len(row) > 0 and row[0]:  # 確保有日期資料
                    print(f"   {i+1}: 日期={row[0]}, 契約={row[1] if len(row)>1 else 'N/A'}, 身份={row[2] if len(row)>2 else 'N/A'}")
            
            # 檢查是否有6/3和6/6的資料
            print(f"\n🔍 搜尋6/3和6/6資料:")
            found_dates = set()
            for row in data[1:]:  # 跳過標題行
                if len(row) > 0:
                    date_str = row[0]
                    if '2025/06/03' in date_str or '2025-06-03' in date_str:
                        found_dates.add('6/3')
                    elif '2025/06/06' in date_str or '2025-06-06' in date_str:
                        found_dates.add('6/6')
            
            if '6/3' in found_dates:
                print("   ✅ 找到6/3資料")
            else:
                print("   ❌ 沒有找到6/3資料")
                
            if '6/6' in found_dates:
                print("   ✅ 找到6/6資料")
            else:
                print("   ❌ 沒有找到6/6資料")
        else:
            print("   ❌ 工作表是空的")
    
    except Exception as e:
        print(f"❌ 檢查歷史資料工作表失敗: {e}")

if __name__ == "__main__":
    main() 