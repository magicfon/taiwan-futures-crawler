#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
檢查Google Sheets中的資料狀況
"""

import sys
import os
import json
from pathlib import Path

# 添加項目根目錄到Python路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from google_sheets_manager import GoogleSheetsManager
    SHEETS_AVAILABLE = True
except ImportError as e:
    print(f"Google Sheets模組導入失敗: {e}")
    SHEETS_AVAILABLE = False

def check_sheets_data():
    """檢查Google Sheets中的資料狀況"""
    print("檢查Google Sheets資料狀況...")
    print("=" * 50)
    
    if not SHEETS_AVAILABLE:
        print("Google Sheets不可用")
        return False
    
    # 初始化Google Sheets管理器
    try:
        sheets_manager = GoogleSheetsManager()
        if not sheets_manager.client:
            print("Google Sheets認證失敗")
            return False
        print("Google Sheets認證成功")
    except Exception as e:
        print(f"初始化失敗: {e}")
        return False
    
    # 連接試算表
    config_file = Path("config/spreadsheet_config.json")
    if config_file.exists():
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
            spreadsheet_id = config.get('spreadsheet_id')
        
        if spreadsheet_id:
            try:
                sheets_manager.connect_spreadsheet(spreadsheet_id)
                print(f"已連接試算表: {sheets_manager.spreadsheet.title}")
                print(f"試算表網址: {sheets_manager.get_spreadsheet_url()}")
            except Exception as e:
                print(f"連接試算表失敗: {e}")
                return False
    else:
        print("找不到試算表配置檔案")
        return False
    
    # 檢查所有工作表的資料
    print("\n檢查各工作表資料...")
    worksheets = sheets_manager.spreadsheet.worksheets()
    
    for worksheet in worksheets:
        print(f"\n工作表: {worksheet.title}")
        try:
            # 獲取所有記錄
            records = worksheet.get_all_records()
            print(f"  資料筆數: {len(records)}")
            
            if len(records) > 0:
                # 顯示前3筆資料的欄位
                first_record = records[0]
                print(f"  欄位數量: {len(first_record)}")
                print("  欄位名稱:")
                for i, field in enumerate(first_record.keys(), 1):
                    print(f"    {i}. {field}")
                
                # 顯示最新的一筆資料
                if '日期' in first_record:
                    latest_date = max(record.get('日期', '') for record in records if record.get('日期'))
                    print(f"  最新日期: {latest_date}")
            else:
                print("  工作表為空")
                
        except Exception as e:
            print(f"  讀取失敗: {e}")
    
    return True

def upload_existing_data():
    """上傳現有的本地資料到Google Sheets"""
    print("\n" + "=" * 50)
    print("檢查本地資料檔案...")
    
    # 檢查output目錄中的檔案
    output_dir = Path("output")
    if not output_dir.exists():
        print("output目錄不存在")
        return False
    
    # 尋找最新的CSV檔案
    csv_files = list(output_dir.glob("*.csv"))
    if not csv_files:
        print("沒有找到CSV檔案")
        return False
    
    # 取得最新的檔案
    latest_csv = max(csv_files, key=lambda x: x.stat().st_mtime)
    print(f"找到最新檔案: {latest_csv.name}")
    
    # 嘗試上傳到Google Sheets
    try:
        import pandas as pd
        
        sheets_manager = GoogleSheetsManager()
        config_file = Path("config/spreadsheet_config.json")
        
        if config_file.exists():
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                spreadsheet_id = config.get('spreadsheet_id')
            
            if spreadsheet_id:
                sheets_manager.connect_spreadsheet(spreadsheet_id)
                
                # 讀取CSV檔案
                df = pd.read_csv(latest_csv, encoding='utf-8-sig')
                print(f"讀取到 {len(df)} 筆資料")
                print(f"欄位: {list(df.columns)}")
                
                # 判斷資料類型
                has_position_fields = any('未平倉' in col for col in df.columns)
                data_type = 'COMPLETE' if has_position_fields else 'TRADING'
                
                print(f"判斷為 {data_type} 資料類型")
                
                # 上傳資料
                result = sheets_manager.upload_data(df, data_type=data_type)
                
                if result:
                    print(f"成功上傳到Google Sheets！")
                    worksheet_name = "完整資料" if data_type == 'COMPLETE' else "交易量資料"
                    print(f"資料已上傳到「{worksheet_name}」工作表")
                else:
                    print("上傳失敗")
                
                return result
    
    except Exception as e:
        print(f"上傳過程發生錯誤: {e}")
        return False

if __name__ == "__main__":
    print("Google Sheets資料檢查工具")
    
    # 檢查現有資料
    check_sheets_data()
    
    # 詢問是否要上傳本地資料
    print("\n" + "=" * 50)
    try:
        response = input("是否要上傳本地資料到Google Sheets? (y/n): ").lower().strip()
        if response in ['y', 'yes', '是']:
            upload_existing_data()
        else:
            print("跳過上傳")
    except:
        # 在某些環境中input可能不可用
        print("自動上傳本地資料...")
        upload_existing_data()
    
    print("\n檢查完成！") 