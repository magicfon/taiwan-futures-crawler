#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
清理Google Sheets並重新同步最近資料
移除舊資料，保留最近30天的資料
"""

from google_sheets_manager import GoogleSheetsManager
from database_manager import TaifexDatabaseManager
import json
from pathlib import Path
import pandas as pd
from datetime import datetime

def main():
    print("🧹 開始清理Google Sheets並重新同步資料...")
    
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
    
    # 檢查現有工作表
    worksheets = sheets_manager.spreadsheet.worksheets()
    worksheet_names = [ws.title for ws in worksheets]
    print(f"📊 現有工作表: {worksheet_names}")
    
    # 1. 清理歷史資料工作表
    print(f"\n🧹 清理歷史資料工作表...")
    try:
        history_ws = sheets_manager.spreadsheet.worksheet('歷史資料')
        
        # 清除所有資料（保留標題行）
        print("   正在清除舊資料...")
        history_ws.batch_clear(["A2:Z50000"])
        print("   ✅ 舊資料已清除")
        
    except Exception as e:
        print(f"   ❌ 清理歷史資料失敗: {e}")
        return
    
    # 2. 刪除不需要的每日摘要相關工作表
    print(f"\n🗑️ 移除不需要的工作表...")
    sheets_to_remove = ['每日摘要', '三大法人趨勢']
    
    for sheet_name in sheets_to_remove:
        try:
            if sheet_name in worksheet_names:
                ws = sheets_manager.spreadsheet.worksheet(sheet_name)
                sheets_manager.spreadsheet.del_worksheet(ws)
                print(f"   ✅ 已刪除: {sheet_name}")
            else:
                print(f"   ⚠️ 工作表不存在: {sheet_name}")
        except Exception as e:
            print(f"   ❌ 刪除工作表 {sheet_name} 失敗: {e}")
    
    # 3. 獲取最近30天的資料
    print(f"\n📊 獲取最近30天的資料...")
    recent_data = db_manager.get_recent_data(days=30)
    print(f"   從資料庫獲取: {len(recent_data)} 筆")
    
    if recent_data.empty:
        print("❌ 沒有資料可以同步")
        return
    
    # 顯示資料範圍
    dates = sorted(recent_data['date'].unique())
    print(f"   📅 日期範圍: {dates[0]} 到 {dates[-1]}")
    print(f"   📅 包含日期: {len(dates)} 天")
    
    # 檢查是否包含6/3和6/6
    target_dates = ['2025/06/03', '2025/06/06']
    found_target_dates = []
    for target_date in target_dates:
        if target_date in dates:
            found_target_dates.append(target_date)
    
    if found_target_dates:
        print(f"   ✅ 找到目標日期: {found_target_dates}")
    else:
        print(f"   ⚠️ 未找到目標日期 {target_dates}")
    
    # 4. 準備Google Sheets格式的資料
    print(f"\n📝 準備上傳資料...")
    sheets_data = []
    for _, row in recent_data.iterrows():
        sheets_row = {
            '日期': row['date'],
            '契約名稱': row['contract_code'],
            '身份別': row['identity_type'],
            '多方交易口數': 0,  # 新資料庫結構中沒有交易口數
            '多方契約金額': 0,
            '空方交易口數': 0,
            '空方契約金額': 0,
            '多空淨額交易口數': 0,
            '多空淨額契約金額': 0,
            '多方未平倉口數': row.get('long_position', 0),
            '多方未平倉契約金額': 0,
            '空方未平倉口數': row.get('short_position', 0),
            '空方未平倉契約金額': 0,
            '多空淨額未平倉口數': row.get('net_position', 0),
            '多空淨額未平倉契約金額': 0,
            '更新時間': row.get('created_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        }
        sheets_data.append(sheets_row)
    
    df = pd.DataFrame(sheets_data)
    print(f"   準備上傳 {len(df)} 筆資料")
    
    # 5. 上傳到Google Sheets
    print(f"\n🚀 開始上傳到歷史資料工作表...")
    try:
        success = sheets_manager.upload_data(df, worksheet_name="歷史資料")
        if success:
            print("✅ 歷史資料同步成功！")
            
            # 6. 驗證上傳結果
            print(f"\n🔍 驗證上傳結果...")
            history_ws = sheets_manager.spreadsheet.worksheet('歷史資料')
            data = history_ws.get_all_values()
            total_rows = len(data)
            print(f"   總行數: {total_rows} 行")
            
            # 檢查目標日期
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
                print("   ⚠️ 未找到6/3和6/6資料")
            
            print(f"\n🎉 清理和同步完成！")
            print(f"🌐 請查看: {sheets_manager.get_spreadsheet_url()}")
            
        else:
            print("❌ 歷史資料同步失敗")
            
    except Exception as e:
        print(f"❌ 上傳過程發生錯誤: {e}")

if __name__ == "__main__":
    main() 