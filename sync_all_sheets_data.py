#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
同步所有Google Sheets工作表資料
確保歷史資料工作表也包含最新資料
"""

import sys
import os
import json
import pandas as pd
from pathlib import Path
from datetime import datetime

# 添加項目根目錄到Python路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from google_sheets_manager import GoogleSheetsManager
    SHEETS_AVAILABLE = True
except ImportError as e:
    print(f"Google Sheets模組導入失敗: {e}")
    SHEETS_AVAILABLE = False

def sync_all_data():
    """同步所有工作表資料"""
    print("同步Google Sheets所有工作表資料")
    print("=" * 60)
    
    if not SHEETS_AVAILABLE:
        print("Google Sheets不可用")
        return False
    
    # 初始化Google Sheets管理器
    try:
        sheets_manager = GoogleSheetsManager()
        config_file = Path("config/spreadsheet_config.json")
        
        if config_file.exists():
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                spreadsheet_id = config.get('spreadsheet_id')
            
            if spreadsheet_id:
                sheets_manager.connect_spreadsheet(spreadsheet_id)
                print(f"已連接試算表: {sheets_manager.spreadsheet.title}")
                print(f"試算表網址: {sheets_manager.get_spreadsheet_url()}")
            else:
                print("找不到spreadsheet_id")
                return False
        else:
            print("找不到試算表配置檔案")
            return False
    except Exception as e:
        print(f"初始化失敗: {e}")
        return False
    
    # 1. 讀取本地最新的完整資料
    print("\n1. 讀取本地最新資料...")
    output_dir = Path("output")
    csv_files = list(output_dir.glob("*.csv"))
    
    if not csv_files:
        print("沒有找到本地CSV檔案")
        return False
    
    latest_csv = max(csv_files, key=lambda x: x.stat().st_mtime)
    print(f"使用檔案: {latest_csv.name}")
    
    try:
        df_local = pd.read_csv(latest_csv, encoding='utf-8-sig')
        print(f"讀取到 {len(df_local)} 筆本地資料")
        print(f"日期範圍: {df_local['日期'].min()} 到 {df_local['日期'].max()}")
    except Exception as e:
        print(f"讀取本地檔案失敗: {e}")
        return False
    
    # 2. 讀取Google Sheets現有的歷史資料
    print("\n2. 讀取Google Sheets歷史資料...")
    try:
        historical_sheet = sheets_manager.spreadsheet.worksheet("歷史資料")
        historical_records = historical_sheet.get_all_records()
        
        if historical_records:
            df_historical = pd.DataFrame(historical_records)
            print(f"現有歷史資料: {len(df_historical)} 筆")
            if '日期' in df_historical.columns:
                print(f"日期範圍: {df_historical['日期'].min()} 到 {df_historical['日期'].max()}")
        else:
            df_historical = pd.DataFrame()
            print("歷史資料工作表為空")
    except Exception as e:
        print(f"讀取歷史資料失敗: {e}")
        df_historical = pd.DataFrame()
    
    # 3. 合併和去重
    print("\n3. 合併和去重資料...")
    try:
        if not df_historical.empty:
            # 合併本地資料和歷史資料
            combined_df = pd.concat([df_historical, df_local], ignore_index=True)
            
            # 去除重複資料（基於日期、契約名稱、身份別）
            combined_df = combined_df.drop_duplicates(
                subset=['日期', '契約名稱', '身份別'], 
                keep='last'  # 保留最新的資料
            )
            
            # 按日期排序
            combined_df = combined_df.sort_values('日期')
            
            print(f"合併後總資料量: {len(combined_df)} 筆")
            print(f"最終日期範圍: {combined_df['日期'].min()} 到 {combined_df['日期'].max()}")
        else:
            combined_df = df_local
            print(f"使用本地資料: {len(combined_df)} 筆")
    except Exception as e:
        print(f"合併資料失敗: {e}")
        return False
    
    # 4. 更新各個工作表
    print("\n4. 更新各個工作表...")
    
    # 判斷資料類型
    has_position_fields = any('未平倉' in col for col in combined_df.columns)
    
    # 4.1 更新歷史資料工作表（完整資料）
    try:
        print("  更新歷史資料工作表...")
        result = sheets_manager.upload_data(combined_df, worksheet_name="歷史資料")
        if result:
            print("  ✅ 歷史資料工作表更新成功")
        else:
            print("  ❌ 歷史資料工作表更新失敗")
    except Exception as e:
        print(f"  ❌ 歷史資料工作表更新錯誤: {e}")
    
    # 4.2 更新完整資料工作表
    if has_position_fields:
        try:
            print("  更新完整資料工作表...")
            result = sheets_manager.upload_data(combined_df, data_type='COMPLETE')
            if result:
                print("  ✅ 完整資料工作表更新成功")
            else:
                print("  ❌ 完整資料工作表更新失敗")
        except Exception as e:
            print(f"  ❌ 完整資料工作表更新錯誤: {e}")
    
    # 4.3 創建交易量資料（移除未平倉欄位）
    if has_position_fields:
        try:
            print("  創建交易量資料...")
            trading_columns = [
                '日期', '契約名稱', '身份別', 
                '多方交易口數', '多方契約金額', 
                '空方交易口數', '空方契約金額',
                '多空淨額交易口數', '多空淨額契約金額'
            ]
            
            # 只保留交易量相關欄位
            trading_df = combined_df[trading_columns].copy()
            
            result = sheets_manager.upload_data(trading_df, data_type='TRADING')
            if result:
                print("  ✅ 交易量資料工作表更新成功")
            else:
                print("  ❌ 交易量資料工作表更新失敗")
        except Exception as e:
            print(f"  ❌ 交易量資料工作表更新錯誤: {e}")
    
    # 5. 最終檢查
    print("\n5. 最終檢查各工作表狀態...")
    worksheets = sheets_manager.spreadsheet.worksheets()
    
    for worksheet in worksheets:
        if worksheet.title in ['歷史資料', '交易量資料', '完整資料']:
            try:
                records = worksheet.get_all_records()
                print(f"  {worksheet.title}: {len(records)} 筆資料")
            except:
                print(f"  {worksheet.title}: 無法讀取")
    
    print("\n" + "=" * 60)
    print("🎉 同步完成！")
    print(f"\n💡 現在可以前往Google試算表查看：")
    print(f"🌐 {sheets_manager.get_spreadsheet_url()}")
    print("\n📊 各工作表說明：")
    print("  - 歷史資料：所有歷史完整資料的匯總")
    print("  - 交易量資料：僅包含交易量的資料（6個數據欄位）")
    print("  - 完整資料：包含交易量和未平倉的完整資料（12個數據欄位）")
    
    return True

if __name__ == "__main__":
    sync_all_data() 