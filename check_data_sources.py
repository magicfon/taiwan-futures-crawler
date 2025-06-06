#!/usr/bin/env python3
"""
檢查所有資料來源，評估歷史資料損失
"""

import sqlite3
import pandas as pd
from pathlib import Path
import os
from datetime import datetime
from google_sheets_manager import GoogleSheetsManager
import json

def main():
    print("🔍 檢查歷史資料損失情況...")
    
    # 1. 檢查本地資料庫
    print("\n📊 檢查本地資料庫...")
    db_path = Path("data/taifex_data.db")
    if db_path.exists():
        conn = sqlite3.connect(db_path)
        
        # 檢查主要資料表
        query = "SELECT MIN(date) as earliest, MAX(date) as latest, COUNT(*) as total FROM futures_data"
        result = pd.read_sql_query(query, conn)
        
        if not result.empty and result.iloc[0]['total'] > 0:
            print(f"✅ 資料庫資料範圍: {result.iloc[0]['earliest']} 到 {result.iloc[0]['latest']}")
            print(f"📈 總筆數: {result.iloc[0]['total']:,} 筆")
            
            # 按月統計
            monthly_query = """
                SELECT substr(date, 1, 7) as month, COUNT(*) as count
                FROM futures_data 
                GROUP BY substr(date, 1, 7) 
                ORDER BY month DESC
            """
            monthly_data = pd.read_sql_query(monthly_query, conn)
            print("📅 按月分布:")
            for _, row in monthly_data.head(10).iterrows():
                print(f"  {row['month']}: {row['count']} 筆")
        else:
            print("❌ 資料庫無資料")
        
        conn.close()
    else:
        print("❌ 找不到本地資料庫")
    
    # 2. 檢查Google Sheets目前狀況
    print("\n☁️ 檢查Google Sheets...")
    try:
        config_file = Path("config/spreadsheet_config.json")
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        sheets_manager = GoogleSheetsManager()
        sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
        
        if sheets_manager.spreadsheet:
            worksheet = sheets_manager.spreadsheet.worksheet("歷史資料")
            all_data = worksheet.get_all_values()
            
            print(f"📊 Google Sheets目前行數: {len(all_data)}")
            
            if len(all_data) > 1:  # 有標題行
                print("📅 Google Sheets資料範圍:")
                # 檢查第2行和最後一行的日期
                if len(all_data) >= 2:
                    first_data = all_data[1]  # 第一筆資料
                    last_data = all_data[-1]  # 最後一筆資料
                    print(f"  最早: {first_data[0] if len(first_data) > 0 else 'N/A'}")
                    print(f"  最新: {last_data[0] if len(last_data) > 0 else 'N/A'}")
            else:
                print("❌ Google Sheets只有標題行，所有資料已被清除")
        
    except Exception as e:
        print(f"❌ 檢查Google Sheets失敗: {e}")
    
    # 3. 檢查本地輸出檔案
    print("\n📁 檢查本地輸出檔案...")
    output_path = Path("output")
    if output_path.exists():
        csv_files = list(output_path.glob("*.csv"))
        excel_files = list(output_path.glob("*.xlsx"))
        
        print(f"📊 找到 {len(csv_files)} 個CSV檔案, {len(excel_files)} 個Excel檔案")
        
        # 列出最近的檔案
        all_files = csv_files + excel_files
        if all_files:
            all_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            print("📅 最近的檔案:")
            for f in all_files[:5]:
                mtime = datetime.fromtimestamp(f.stat().st_mtime)
                print(f"  {f.name} - {mtime.strftime('%Y-%m-%d %H:%M:%S')}")
    else:
        print("❌ 找不到輸出目錄")
    
    # 4. 恢復建議
    print("\n🔧 恢復建議:")
    print("1. 本地資料庫是主要資料來源，可以從這裡恢復")
    print("2. 建議立即備份現有資料庫")
    print("3. 修改Google Sheets上傳策略，避免再次清除歷史資料")
    print("4. 可以重新上傳所有歷史資料到Google Sheets")

if __name__ == "__main__":
    main() 