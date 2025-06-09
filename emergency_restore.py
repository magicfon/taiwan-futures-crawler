#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
緊急資料恢復和問題診斷
"""

from database_manager import TaifexDatabaseManager
from google_sheets_manager import GoogleSheetsManager
import json
import pandas as pd
import sqlite3
from pathlib import Path

def main():
    print("🚨 緊急資料恢復和問題診斷...")
    
    # 初始化管理器
    db_manager = TaifexDatabaseManager()
    db_path = db_manager.db_path
    
    # 1. 檢查資料庫中的重複資料
    print("\n🔍 檢查資料庫重複資料...")
    conn = sqlite3.connect(db_path)
    
    query = """
    SELECT date, contract_code, identity_type, COUNT(*) as count
    FROM futures_data 
    GROUP BY date, contract_code, identity_type 
    HAVING COUNT(*) > 1
    ORDER BY date DESC, count DESC
    """
    
    duplicates = pd.read_sql_query(query, conn)
    if not duplicates.empty:
        print(f"❌ 發現 {len(duplicates)} 組重複資料:")
        for _, row in duplicates.head(10).iterrows():
            print(f"   - {row['date']} {row['contract_code']} {row['identity_type']}: {row['count']} 筆")
    else:
        print("✅ 沒有發現重複資料")
    
    # 2. 檢查最新的正確資料
    print("\n📊 檢查最新的正確資料...")
    query = """
    SELECT DISTINCT date, contract_code, identity_type, 
           long_position, short_position, net_position
    FROM futures_data 
    WHERE date >= '2025/06/02'
    ORDER BY date DESC, contract_code, identity_type
    """
    
    recent_data = pd.read_sql_query(query, conn)
    conn.close()
    
    print(f"📈 最近資料: {len(recent_data)} 筆")
    
    if not recent_data.empty:
        # 按日期分組顯示
        for date in sorted(recent_data['date'].unique(), reverse=True)[:5]:
            date_data = recent_data[recent_data['date'] == date]
            print(f"\n📅 {date}: {len(date_data)} 筆資料")
            for _, row in date_data.iterrows():
                print(f"   - {row['contract_code']} {row['identity_type']}: 多{row['long_position']}, 空{row['short_position']}, 淨{row['net_position']}")
    
    # 3. 檢查是否有完整的6/2-6/6資料
    print("\n🔍 檢查6/2-6/6完整資料...")
    target_dates = ['2025/06/02', '2025/06/03', '2025/06/04', '2025/06/05', '2025/06/06']
    
    for date in target_dates:
        date_data = recent_data[recent_data['date'] == date]
        if len(date_data) > 0:
            contracts = date_data['contract_code'].unique()
            identities = date_data['identity_type'].unique()
            print(f"   ✅ {date}: {len(date_data)}筆 (契約: {contracts}, 身份: {identities})")
        else:
            print(f"   ❌ {date}: 無資料")
    
    # 4. 提供修復選項
    print(f"\n🔧 修復選項:")
    print(f"   1. 清除資料庫重複資料")
    print(f"   2. 重新從台期所爬取6/2-6/6資料") 
    print(f"   3. 只清理Google Sheets重複上傳")
    
    user_choice = input("\n請選擇修復方案 (1/2/3): ").strip()
    
    if user_choice == "1":
        print("\n🧹 開始清除資料庫重複資料...")
        # 保留最新的記錄，刪除較舊的重複項
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        cleanup_query = """
        DELETE FROM futures_data 
        WHERE id NOT IN (
            SELECT MAX(id) 
            FROM futures_data 
            GROUP BY date, contract_code, identity_type
        )
        """
        
        try:
            cursor.execute(cleanup_query)
            deleted_count = cursor.rowcount
            conn.commit()
            print(f"✅ 資料庫重複資料清除完成，刪除了 {deleted_count} 筆重複資料")
        except Exception as e:
            print(f"❌ 清除失敗: {e}")
            conn.rollback()
        finally:
            conn.close()
    
    elif user_choice == "2":
        print("\n🕷️ 重新爬取6/2-6/6資料...")
        import subprocess
        
        # 轉換為爬蟲需要的日期格式
        crawler_dates = ['2025-06-02', '2025-06-03', '2025-06-04', '2025-06-05', '2025-06-06']
        
        for date in crawler_dates:
            try:
                result = subprocess.run([
                    'python', 'taifex_crawler.py', '--date', date
                ], capture_output=True, text=True, encoding='utf-8')
                
                if result.returncode == 0:
                    print(f"   ✅ {date} 爬取成功")
                else:
                    print(f"   ❌ {date} 爬取失敗: {result.stderr}")
            except Exception as e:
                print(f"   ❌ {date} 爬取錯誤: {e}")
    
    elif user_choice == "3":
        print("\n🧹 只清理Google Sheets...")
        sheets_manager = GoogleSheetsManager()
        
        # 載入配置
        config_file = Path('config/spreadsheet_config.json')
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
        
        if sheets_manager.spreadsheet:
            try:
                history_ws = sheets_manager.spreadsheet.worksheet('歷史資料')
                history_ws.batch_clear(["A2:Z50000"])
                print("✅ Google Sheets 歷史資料已清空")
                
                # 重新上傳正確的資料
                if not recent_data.empty:
                    print("📤 重新上傳正確資料...")
                    # 轉換為Google Sheets格式
                    sheets_data = []
                    for _, row in recent_data.iterrows():
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
                    success = sheets_manager.upload_data(df, worksheet_name="歷史資料")
                    if success:
                        print("✅ 正確資料已重新上傳")
                    else:
                        print("❌ 重新上傳失敗")
                        
            except Exception as e:
                print(f"❌ Google Sheets 清理失敗: {e}")

if __name__ == "__main__":
    main() 