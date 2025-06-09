#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
修復資料格式問題 - 合併多空資料
"""

import sqlite3
import pandas as pd
from database_manager import TaifexDatabaseManager
from google_sheets_manager import GoogleSheetsManager
import json
from pathlib import Path

def main():
    print("🔧 修復資料格式問題...")
    
    # 連接資料庫
    db_manager = TaifexDatabaseManager()
    db_path = db_manager.db_path
    conn = sqlite3.connect(db_path)
    
    # 1. 檢查現在的資料結構
    print("\n📊 檢查目前資料結構...")
    query = """
    SELECT date, contract_code, identity_type, 
           SUM(CASE WHEN long_position > 0 THEN long_position ELSE 0 END) as total_long,
           SUM(CASE WHEN short_position > 0 THEN short_position ELSE 0 END) as total_short,
           SUM(net_position) as total_net
    FROM futures_data 
    WHERE date >= '2025/06/02'
    GROUP BY date, contract_code, identity_type
    ORDER BY date DESC, contract_code, identity_type
    """
    
    clean_data = pd.read_sql_query(query, conn)
    conn.close()
    
    print(f"✅ 清理後資料: {len(clean_data)} 筆")
    
    # 2. 顯示修復後的資料
    if not clean_data.empty:
        print(f"\n📋 修復後資料預覽:")
        for date in sorted(clean_data['date'].unique(), reverse=True)[:3]:
            date_data = clean_data[clean_data['date'] == date]
            print(f"\n📅 {date}: {len(date_data)} 筆資料")
            for _, row in date_data.iterrows():
                print(f"   - {row['contract_code']} {row['identity_type']}: 多{row['total_long']}, 空{row['total_short']}, 淨{row['total_net']}")
    
    # 3. 重新上傳到Google Sheets
    print(f"\n🚀 重新上傳到Google Sheets...")
    
    sheets_manager = GoogleSheetsManager()
    config_file = Path('config/spreadsheet_config.json')
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
    
    if sheets_manager.spreadsheet:
        try:
            # 清空歷史資料
            history_ws = sheets_manager.spreadsheet.worksheet('歷史資料')
            history_ws.batch_clear(["A2:Z50000"])
            print("✅ Google Sheets 歷史資料已清空")
            
            # 準備正確格式的資料
            sheets_data = []
            for _, row in clean_data.iterrows():
                sheets_row = {
                    '日期': row['date'],
                    '契約名稱': row['contract_code'],
                    '身份別': row['identity_type'],
                    '多方交易口數': 0,
                    '多方契約金額': 0,
                    '空方交易口數': 0,
                    '空方契約金額': 0,
                    '多空淨額交易口數': 0,
                    '多空淨額契約金額': 0,
                    '多方未平倉口數': row['total_long'],
                    '多方未平倉契約金額': 0,
                    '空方未平倉口數': row['total_short'],
                    '空方未平倉契約金額': 0,
                    '多空淨額未平倉口數': row['total_net'],
                    '多空淨額未平倉契約金額': 0,
                    '更新時間': '',
                }
                sheets_data.append(sheets_row)
            
            df = pd.DataFrame(sheets_data)
            print(f"📤 準備上傳 {len(df)} 筆正確格式資料...")
            
            success = sheets_manager.upload_data(df, worksheet_name="歷史資料")
            if success:
                print("✅ 正確格式資料已上傳到Google Sheets！")
                print(f"🌐 請查看: {sheets_manager.get_spreadsheet_url()}")
                
                # 驗證結果
                print(f"\n🔍 驗證上傳結果...")
                data = history_ws.get_all_values()
                print(f"   總行數: {len(data)} 行 (包含標題)")
                
                # 檢查是否還有重複
                date_contract_identity = []
                for row in data[1:]:  # 跳過標題
                    if len(row) >= 3:
                        key = f"{row[0]}_{row[1]}_{row[2]}"
                        date_contract_identity.append(key)
                
                duplicates = len(date_contract_identity) - len(set(date_contract_identity))
                if duplicates == 0:
                    print("   ✅ 沒有重複資料")
                else:
                    print(f"   ⚠️ 仍有 {duplicates} 筆重複資料")
                
                # 檢查6/3和6/6資料
                found_dates = set()
                for row in data[1:]:
                    if len(row) > 0:
                        date_str = row[0]
                        if '2025/06/03' in date_str:
                            found_dates.add('6/3')
                        elif '2025/06/06' in date_str:
                            found_dates.add('6/6')
                
                if '6/3' in found_dates:
                    print("   ✅ 確認找到6/3資料")
                if '6/6' in found_dates:
                    print("   ✅ 確認找到6/6資料")
                    
            else:
                print("❌ 上傳失敗")
                
        except Exception as e:
            print(f"❌ 處理失敗: {e}")

if __name__ == "__main__":
    main() 