#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
測試Google Sheets兩種資料類型上傳功能
"""

import sys
import os
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

# 添加項目根目錄到Python路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from google_sheets_manager import GoogleSheetsManager
    SHEETS_AVAILABLE = True
    print("✅ Google Sheets模組載入成功")
except ImportError as e:
    print(f"❌ Google Sheets模組導入失敗: {e}")
    SHEETS_AVAILABLE = False

def test_two_data_types():
    """測試兩種資料類型的上傳功能"""
    print("🧪 測試Google Sheets兩種資料類型上傳")
    print("=" * 60)
    print(f"⏰ 測試時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    if not SHEETS_AVAILABLE:
        print("❌ Google Sheets不可用，無法進行測試")
        return False
    
    # 1. 初始化Google Sheets管理器
    print("🔧 初始化Google Sheets管理器...")
    try:
        sheets_manager = GoogleSheetsManager()
        if not sheets_manager.client:
            print("❌ Google Sheets認證失敗")
            return False
        print("✅ Google Sheets認證成功")
    except Exception as e:
        print(f"❌ 初始化失敗: {e}")
        return False
    
    # 2. 連接試算表
    print("\n📊 連接Google試算表...")
    config_file = Path("config/spreadsheet_config.json")
    
    if config_file.exists():
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
            spreadsheet_id = config.get('spreadsheet_id')
        
        if spreadsheet_id:
            try:
                sheets_manager.connect_spreadsheet(spreadsheet_id)
                print(f"✅ 已連接試算表: {sheets_manager.spreadsheet.title}")
                print(f"🌐 試算表網址: {sheets_manager.get_spreadsheet_url()}")
            except Exception as e:
                print(f"❌ 連接試算表失敗: {e}")
                return False
        else:
            print("❌ 配置檔案中找不到spreadsheet_id")
            return False
    else:
        print("❌ 找不到試算表配置檔案")
        return False
    
    # 3. 檢查現有工作表
    print("\n📋 檢查現有工作表...")
    worksheets = sheets_manager.spreadsheet.worksheets()
    worksheet_names = [ws.title for ws in worksheets]
    print(f"現有工作表: {', '.join(worksheet_names)}")
    
    # 4. 創建測試資料
    print("\n📊 創建測試資料...")
    
    # TRADING模式測試資料（6個數據欄位）
    trading_data = pd.DataFrame({
        '日期': ['2025/06/09'],
        '契約名稱': ['TX'],
        '身份別': ['測試用戶'],
        '多方交易口數': [1000],
        '多方契約金額': [50000000],
        '空方交易口數': [800],
        '空方契約金額': [40000000],
        '多空淨額交易口數': [200],
        '多空淨額契約金額': [10000000]
    })
    print(f"✅ TRADING測試資料已創建 ({len(trading_data.columns)}個欄位)")
    
    # COMPLETE模式測試資料（12個數據欄位）
    complete_data = pd.DataFrame({
        '日期': ['2025/06/09'],
        '契約名稱': ['TX'],
        '身份別': ['測試用戶'],
        '多方交易口數': [1000],
        '多方契約金額': [50000000],
        '空方交易口數': [800],
        '空方契約金額': [40000000],
        '多空淨額交易口數': [200],
        '多空淨額契約金額': [10000000],
        '多方未平倉口數': [5000],
        '多方未平倉契約金額': [250000000],
        '空方未平倉口數': [4500],
        '空方未平倉契約金額': [225000000],
        '多空淨額未平倉口數': [500],
        '多空淨額未平倉契約金額': [25000000]
    })
    print(f"✅ COMPLETE測試資料已創建 ({len(complete_data.columns)}個欄位)")
    
    # 5. 測試TRADING模式上傳
    print("\n🚀 測試TRADING模式上傳...")
    try:
        result = sheets_manager.upload_data(trading_data, data_type='TRADING')
        if result:
            print("✅ TRADING模式上傳成功 -> 應該出現在「交易量資料」工作表")
        else:
            print("❌ TRADING模式上傳失敗")
    except Exception as e:
        print(f"❌ TRADING模式上傳錯誤: {e}")
    
    # 6. 測試COMPLETE模式上傳
    print("\n🚀 測試COMPLETE模式上傳...")
    try:
        result = sheets_manager.upload_data(complete_data, data_type='COMPLETE')
        if result:
            print("✅ COMPLETE模式上傳成功 -> 應該出現在「完整資料」工作表")
        else:
            print("❌ COMPLETE模式上傳失敗")
    except Exception as e:
        print(f"❌ COMPLETE模式上傳錯誤: {e}")
    
    # 7. 測試自動判斷功能
    print("\n🤖 測試自動判斷功能...")
    try:
        # 使用complete_data但不指定data_type，應該自動判斷為COMPLETE
        result = sheets_manager.upload_data(complete_data, data_type=None)
        if result:
            print("✅ 自動判斷上傳成功 -> 應該自動判斷為完整資料")
        else:
            print("❌ 自動判斷上傳失敗")
    except Exception as e:
        print(f"❌ 自動判斷上傳錯誤: {e}")
    
    # 8. 檢查最終工作表狀態
    print("\n📊 檢查最終工作表狀態...")
    worksheets = sheets_manager.spreadsheet.worksheets()
    worksheet_names = [ws.title for ws in worksheets]
    
    trading_exists = "交易量資料" in worksheet_names
    complete_exists = "完整資料" in worksheet_names
    
    print(f"交易量資料工作表: {'✅ 存在' if trading_exists else '❌ 不存在'}")
    print(f"完整資料工作表: {'✅ 存在' if complete_exists else '❌ 不存在'}")
    
    # 9. 顯示各工作表的資料筆數
    if trading_exists:
        try:
            trading_sheet = sheets_manager.spreadsheet.worksheet("交易量資料")
            trading_records = trading_sheet.get_all_records()
            print(f"交易量資料筆數: {len(trading_records)}")
        except Exception as e:
            print(f"無法讀取交易量資料: {e}")
    
    if complete_exists:
        try:
            complete_sheet = sheets_manager.spreadsheet.worksheet("完整資料")
            complete_records = complete_sheet.get_all_records()
            print(f"完整資料筆數: {len(complete_records)}")
        except Exception as e:
            print(f"無法讀取完整資料: {e}")
    
    print("\n" + "=" * 60)
    print("🎉 測試完成！")
    
    # 10. 總結
    print("\n💡 測試總結:")
    print(f"  - Google試算表網址: {sheets_manager.get_spreadsheet_url()}")
    print("  - TRADING模式會上傳到「交易量資料」工作表")
    print("  - COMPLETE模式會上傳到「完整資料」工作表")
    print("  - 支援根據資料欄位自動判斷資料類型")
    print("  - 在主爬蟲程式中會根據--data_type參數自動選擇")
    
    return True

if __name__ == "__main__":
    test_two_data_types() 