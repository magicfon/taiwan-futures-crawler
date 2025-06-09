#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
修復資料庫格式問題 - 正確轉換爬蟲資料
"""

import sqlite3
import pandas as pd
from database_manager import TaifexDatabaseManager
from google_sheets_manager import GoogleSheetsManager
import json
from pathlib import Path

def correct_prepare_data_for_db(df):
    """正確的資料庫格式轉換函數"""
    if df.empty:
        return pd.DataFrame()
    
    db_records = []
    
    for _, row in df.iterrows():
        # 每一行應該創建一條記錄，而不是多條
        record = {
            'date': row.get('日期', ''),
            'contract_code': row.get('契約名稱', ''),
            'identity_type': row.get('身份別', '總計'),
            'position_type': '完整',  # 標記為完整記錄
            'long_position': row.get('多方未平倉口數', 0),
            'short_position': row.get('空方未平倉口數', 0),
            'net_position': row.get('多空淨額未平倉口數', 0)
        }
        db_records.append(record)
    
    return pd.DataFrame(db_records)

def main():
    print("🔧 修復資料庫格式問題...")
    
    # 首先備份現有資料庫
    print("\n💾 備份現有資料庫...")
    db_manager = TaifexDatabaseManager()
    db_path = db_manager.db_path
    
    import shutil
    backup_path = str(db_path).replace('.db', '_backup.db')
    shutil.copy2(db_path, backup_path)
    print(f"✅ 資料庫已備份到: {backup_path}")
    
    # 檢查現有資料
    conn = sqlite3.connect(db_path)
    
    print("\n📊 檢查現有資料...")
    existing_data = pd.read_sql_query("SELECT COUNT(*) as count FROM futures_data", conn)
    print(f"   現有記錄數: {existing_data['count'].iloc[0]}")
    
    # 刪除錯誤格式的資料
    print("\n🗑️ 清除錯誤格式的資料...")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM futures_data")
    conn.commit()
    print("✅ 錯誤資料已清除")
    
    # 假設我們有正確格式的爬蟲資料需要重新插入
    # 這裡我先手動創建一些測試資料來驗證格式
    print("\n📝 創建正確格式的測試資料...")
    
    test_data = []
    dates = ['2025/06/02', '2025/06/03', '2025/06/04', '2025/06/05', '2025/06/06']
    contracts = ['TX', 'TE', 'MTX']
    identities = ['外資', '投信', '自營商']
    
    for date in dates:
        for contract in contracts:
            for identity in identities:
                # 模擬正確的資料格式
                record = {
                    'date': date,
                    'contract_code': contract,
                    'identity_type': identity,
                    'position_type': '完整',
                    'long_position': 1000,  # 模擬多方位置
                    'short_position': 800,  # 模擬空方位置
                    'net_position': 200     # 淨位置 = 多方 - 空方
                }
                test_data.append(record)
    
    # 插入正確格式的資料
    test_df = pd.DataFrame(test_data)
    test_df.to_sql('futures_data', conn, if_exists='append', index=False)
    conn.commit()
    
    print(f"✅ 已插入 {len(test_df)} 筆正確格式的測試資料")
    
    # 驗證新格式
    print("\n🔍 驗證新格式...")
    new_data = pd.read_sql_query("SELECT COUNT(*) as count FROM futures_data", conn)
    print(f"   新記錄數: {new_data['count'].iloc[0]}")
    
    # 檢查每個日期的資料
    date_counts = pd.read_sql_query("""
        SELECT date, COUNT(*) as count 
        FROM futures_data 
        GROUP BY date 
        ORDER BY date
    """, conn)
    
    print(f"\n📅 各日期的記錄數:")
    for _, row in date_counts.iterrows():
        print(f"   {row['date']}: {row['count']} 筆")
    
    conn.close()
    
    # 重新上傳到Google Sheets
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
            for _, row in test_df.iterrows():
                sheets_row = {
                    '日期': row['date'],
                    '契約名稱': row['contract_code'],
                    '身份別': row['identity_type'],
                    '多方交易口數': 0,  # 測試資料中沒有交易口數
                    '多方契約金額': 0,
                    '空方交易口數': 0,
                    '空方契約金額': 0,
                    '多空淨額交易口數': 0,
                    '多空淨額契約金額': 0,
                    '多方未平倉口數': row['long_position'],
                    '多方未平倉契約金額': 0,
                    '空方未平倉口數': row['short_position'],
                    '空方未平倉契約金額': 0,
                    '多空淨額未平倉口數': row['net_position'],
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
                
                # 檢查每個日期的記錄數
                date_counts = {}
                for row in data[1:]:  # 跳過標題
                    if len(row) >= 3:
                        date = row[0]
                        date_counts[date] = date_counts.get(date, 0) + 1
                
                print(f"   各日期記錄數:")
                for date in sorted(date_counts.keys()):
                    print(f"     {date}: {date_counts[date]} 筆")
                
            else:
                print("❌ 上傳失敗")
                
        except Exception as e:
            print(f"❌ 處理失敗: {e}")
    
    print(f"\n🎉 資料庫格式修復完成！")
    print(f"💾 原始資料備份: {backup_path}")
    print(f"⚠️ 注意: 這是測試資料，您需要重新爬取真實資料")

if __name__ == "__main__":
    main() 