#!/usr/bin/env python3
"""
歷史資料恢復工具
從多個來源恢復歷史資料到Google Sheets
"""

import pandas as pd
import json
from pathlib import Path
from datetime import datetime, timedelta
from google_sheets_manager import GoogleSheetsManager
from database_manager import TaifexDatabaseManager
import sqlite3

def restore_from_output_files():
    """從output目錄的CSV/Excel檔案恢復資料"""
    print("📁 從輸出檔案恢復資料...")
    
    output_path = Path("output")
    all_data = []
    
    if output_path.exists():
        # 尋找所有CSV檔案
        csv_files = list(output_path.glob("taifex_*.csv"))
        csv_files.sort(key=lambda x: x.name)  # 按檔名排序
        
        for csv_file in csv_files:
            try:
                df = pd.read_csv(csv_file, encoding='utf-8-sig')
                if not df.empty:
                    print(f"  ✅ 讀取 {csv_file.name}: {len(df)} 筆")
                    all_data.append(df)
            except Exception as e:
                print(f"  ❌ 讀取 {csv_file.name} 失敗: {e}")
    
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        # 去重並排序
        combined_df = combined_df.drop_duplicates()
        if '日期' in combined_df.columns:
            combined_df = combined_df.sort_values('日期')
        print(f"✅ 從輸出檔案恢復 {len(combined_df)} 筆資料")
        return combined_df
    
    return pd.DataFrame()

def restore_from_database():
    """從資料庫恢復所有資料"""
    print("🗄️ 從資料庫恢復資料...")
    
    try:
        db_manager = TaifexDatabaseManager()
        # 獲取所有歷史資料
        query = """
            SELECT date, contract_code, identity_type, 
                   long_position, short_position, net_position,
                   created_at
            FROM futures_data 
            ORDER BY date, contract_code, identity_type
        """
        
        db_path = Path("data/taifex_data.db")
        if db_path.exists():
            conn = sqlite3.connect(db_path)
            df = pd.read_sql_query(query, conn)
            conn.close()
            
            if not df.empty:
                # 轉換為Google Sheets格式
                converted_data = []
                for _, row in df.iterrows():
                    converted_row = {
                        '日期': row.get('date', ''),
                        '契約名稱': row.get('contract_code', ''),
                        '身份別': row.get('identity_type', ''),
                        '多方交易口數': row.get('long_position', 0),
                        '多方契約金額': '',
                        '空方交易口數': row.get('short_position', 0),
                        '空方契約金額': '',
                        '多空淨額交易口數': row.get('net_position', 0),
                        '多空淨額契約金額': '',
                        '多方未平倉口數': '',
                        '多方未平倉契約金額': '',
                        '空方未平倉口數': '',
                        '空方未平倉契約金額': '',
                        '多空淨額未平倉口數': '',
                        '多空淨額未平倉契約金額': '',
                        '更新時間': row.get('created_at', '')
                    }
                    converted_data.append(converted_row)
                
                result_df = pd.DataFrame(converted_data)
                print(f"✅ 從資料庫恢復 {len(result_df)} 筆資料")
                return result_df
        
    except Exception as e:
        print(f"❌ 從資料庫恢復失敗: {e}")
    
    return pd.DataFrame()

def check_google_sheets_backup():
    """檢查Google Sheets是否有其他工作表可以恢復"""
    print("☁️ 檢查Google Sheets其他工作表...")
    
    try:
        config_file = Path("config/spreadsheet_config.json")
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        sheets_manager = GoogleSheetsManager()
        sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
        
        if sheets_manager.spreadsheet:
            worksheets = sheets_manager.spreadsheet.worksheets()
            
            print("📋 可用的工作表:")
            for ws in worksheets:
                row_count = len(ws.get_all_values())
                print(f"  - {ws.title}: {row_count} 行")
                
                # 如果找到有資料的工作表（除了歷史資料）
                if ws.title != "歷史資料" and row_count > 1:
                    print(f"    💡 發現可能的備份資料：{ws.title}")
    
    except Exception as e:
        print(f"❌ 檢查Google Sheets失敗: {e}")

def main():
    print("🔧 歷史資料恢復工具")
    print("=" * 50)
    
    # 1. 檢查各種恢復來源
    print("\n🔍 檢查恢復來源...")
    check_google_sheets_backup()
    
    # 2. 從資料庫恢復
    db_data = restore_from_database()
    
    # 3. 從輸出檔案恢復
    file_data = restore_from_output_files()
    
    # 4. 合併所有資料
    print("\n🔄 合併恢復的資料...")
    all_restored_data = []
    
    if not db_data.empty:
        all_restored_data.append(db_data)
        print(f"  ✅ 資料庫: {len(db_data)} 筆")
    
    if not file_data.empty:
        all_restored_data.append(file_data)
        print(f"  ✅ 輸出檔案: {len(file_data)} 筆")
    
    if all_restored_data:
        final_data = pd.concat(all_restored_data, ignore_index=True)
        final_data = final_data.drop_duplicates()
        
        if '日期' in final_data.columns:
            final_data = final_data.sort_values('日期')
        
        print(f"\n📊 總共恢復 {len(final_data)} 筆資料")
        
        # 5. 詢問是否要恢復到Google Sheets
        print("\n⚠️  重要提醒：")
        print("1. 此操作將恢復歷史資料到Google Sheets")
        print("2. 如果資料量過大，可能會接近Google Sheets限制")
        print("3. 建議先備份目前的Google Sheets內容")
        
        response = input("\n是否要恢復到Google Sheets？(y/N): ")
        
        if response.lower() == 'y':
            print("\n⬆️ 上傳恢復的資料到Google Sheets...")
            
            config_file = Path("config/spreadsheet_config.json")
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            sheets_manager = GoogleSheetsManager()
            sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
            
            if sheets_manager.upload_data(final_data):
                print("✅ 歷史資料恢復完成！")
            else:
                print("❌ 恢復失敗")
        else:
            # 保存到本地檔案
            output_file = f"backup/restored_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            final_data.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f"💾 恢復的資料已保存到: {output_file}")
    
    else:
        print("\n❌ 沒有找到可恢復的歷史資料")

if __name__ == "__main__":
    main() 