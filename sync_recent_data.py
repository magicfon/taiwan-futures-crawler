#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
同步最近資料到Google Sheets歷史資料工作表
"""

from google_sheets_manager import GoogleSheetsManager
from database_manager import TaifexDatabaseManager
import json
from pathlib import Path
import pandas as pd

def main():
    print("🚀 開始同步最近資料到Google Sheets歷史資料工作表...")
    
    # 初始化管理器
    db_manager = TaifexDatabaseManager()
    sheets_manager = GoogleSheetsManager()
    
    if not sheets_manager.client:
        print("❌ Google Sheets認證失敗")
        return
    
    # 連接到試算表
    config_file = Path('config/spreadsheet_config.json')
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
    if not sheets_manager.spreadsheet:
        print("❌ 無法連接到試算表")
        return
    
    print(f"✅ 已連接到試算表: {sheets_manager.spreadsheet.title}")
    
    # 獲取最近7天的資料庫資料
    recent_data = db_manager.get_recent_data(days=7)
    print(f"📊 從資料庫獲取最近7天資料: {len(recent_data)} 筆")
    
    if recent_data.empty:
        print("❌ 沒有資料可以同步")
        return
    
    # 顯示資料範圍
    dates = recent_data['date'].unique()
    print(f"📅 日期範圍: {dates}")
    
    # 準備Google Sheets格式的資料
    sheets_data = []
    for _, row in recent_data.iterrows():
        # 轉換資料格式以符合Google Sheets
        sheets_row = {
            '日期': row['date'],
            '契約名稱': row['contract_code'],
            '身份別': row['identity_type'],
            '多方交易口數': row.get('long_trade_volume', 0),
            '多方契約金額': row.get('long_trade_amount', 0),
            '空方交易口數': row.get('short_trade_volume', 0),
            '空方契約金額': row.get('short_trade_amount', 0),
            '多空淨額交易口數': row.get('net_trade_volume', 0),
            '多空淨額契約金額': row.get('net_trade_amount', 0),
            '多方未平倉口數': row.get('long_position_volume', 0),
            '多方未平倉契約金額': row.get('long_position_amount', 0),
            '空方未平倉口數': row.get('short_position_volume', 0),
            '空方未平倉契約金額': row.get('short_position_amount', 0),
            '多空淨額未平倉口數': row.get('net_position_volume', 0),
            '多空淨額未平倉契約金額': row.get('net_position_amount', 0),
        }
        sheets_data.append(sheets_row)
    
    # 轉換為DataFrame
    df = pd.DataFrame(sheets_data)
    print(f"📝 準備上傳 {len(df)} 筆資料")
    
    # 顯示將要上傳的資料概覽
    print(f"\n📋 將要上傳的資料:")
    for date in sorted(df['日期'].unique()):
        date_data = df[df['日期'] == date]
        print(f"   📅 {date}: {len(date_data)} 筆資料")
        for _, row in date_data.iterrows():
            print(f"      - {row['契約名稱']} {row['身份別']}: 多方{row['多方未平倉口數']}口, 空方{row['空方未平倉口數']}口")
    
    # 上傳到Google Sheets
    print(f"\n🚀 開始上傳到歷史資料工作表...")
    try:
        success = sheets_manager.upload_data(df, worksheet_name="歷史資料")
        if success:
            print("✅ 歷史資料同步成功！")
            print(f"🌐 請查看: {sheets_manager.get_spreadsheet_url()}")
            
            # 再次檢查是否有6/3和6/6資料
            print(f"\n🔍 檢查上傳結果...")
            worksheet = sheets_manager.spreadsheet.worksheet('歷史資料')
            data = worksheet.get_all_values()
            
            found_dates = set()
            for row in data[1:]:  # 跳過標題行
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
            
            if not found_dates:
                print("   ⚠️ 仍然沒有找到6/3和6/6資料，可能需要檢查資料格式")
        else:
            print("❌ 歷史資料同步失敗")
    except Exception as e:
        print(f"❌ 上傳過程發生錯誤: {e}")

if __name__ == "__main__":
    main() 