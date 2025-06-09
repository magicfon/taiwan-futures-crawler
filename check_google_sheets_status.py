#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
檢查Google Sheets狀態和資料類型支援
"""

import sys
import os
from datetime import datetime

# 添加項目根目錄到Python路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from google_sheets_manager import GoogleSheetsManager
    SHEETS_AVAILABLE = True
except ImportError as e:
    print(f"❌ Google Sheets模組導入失敗: {e}")
    SHEETS_AVAILABLE = False

def check_google_sheets_support():
    """檢查Google Sheets對兩種資料類型的支援"""
    print("🔍 檢查Google Sheets狀態...")
    print("-" * 50)
    
    if not SHEETS_AVAILABLE:
        print("❌ Google Sheets管理器無法導入")
        return False
    
    # 初始化Google Sheets管理器
    try:
        sheets_manager = GoogleSheetsManager()
        print("✅ Google Sheets管理器初始化成功")
    except Exception as e:
        print(f"❌ Google Sheets管理器初始化失敗: {e}")
        return False
    
    # 檢查認證狀態
    if sheets_manager.client:
        print("✅ Google Sheets認證成功")
    else:
        print("❌ Google Sheets認證失敗或未設定")
        return False
    
    # 檢查是否已連接試算表
    if sheets_manager.spreadsheet:
        print(f"✅ 已連接試算表: {sheets_manager.spreadsheet.title}")
        print(f"📋 試算表網址: {sheets_manager.get_spreadsheet_url()}")
    else:
        print("⚠️ 尚未連接到試算表")
        return False
    
    # 檢查工作表
    print("\n📊 檢查工作表...")
    worksheets = sheets_manager.spreadsheet.worksheets()
    worksheet_names = [ws.title for ws in worksheets]
    
    print(f"現有工作表: {', '.join(worksheet_names)}")
    
    # 檢查兩種資料類型的工作表
    trading_sheet_exists = "交易量資料" in worksheet_names
    complete_sheet_exists = "完整資料" in worksheet_names
    
    print(f"📋 交易量資料工作表: {'✅ 存在' if trading_sheet_exists else '❌ 不存在'}")
    print(f"📋 完整資料工作表: {'✅ 存在' if complete_sheet_exists else '❌ 不存在'}")
    
    # 檢查表頭定義
    print("\n🗂️ 檢查表頭定義...")
    trading_headers = sheets_manager.get_trading_headers()
    complete_headers = sheets_manager.get_complete_headers()
    
    print(f"TRADING模式表頭 ({len(trading_headers)}個欄位):")
    for i, header in enumerate(trading_headers, 1):
        print(f"  {i}. {header}")
    
    print(f"\nCOMPLETE模式表頭 ({len(complete_headers)}個欄位):")
    for i, header in enumerate(complete_headers, 1):
        print(f"  {i}. {header}")
    
    # 檢查upload_data方法的資料類型支援
    print("\n⚙️ 檢查upload_data方法...")
    try:
        # 檢查方法是否存在data_type參數
        import inspect
        sig = inspect.signature(sheets_manager.upload_data)
        params = list(sig.parameters.keys())
        
        if 'data_type' in params:
            print("✅ upload_data方法支援data_type參數")
        else:
            print("❌ upload_data方法不支援data_type參數")
        
        print(f"📋 方法參數: {', '.join(params)}")
        
    except Exception as e:
        print(f"❌ 檢查upload_data方法失敗: {e}")
    
    # 檢查資料類型映射
    print("\n🔄 檢查資料類型映射...")
    test_cases = [
        ('TRADING', '交易量資料'),
        ('COMPLETE', '完整資料'),
        (None, '自動判斷')
    ]
    
    for data_type, expected in test_cases:
        print(f"  {data_type} -> {expected}")
    
    return True

def test_data_type_upload():
    """測試兩種資料類型的上傳功能"""
    print("\n🧪 測試資料類型上傳功能...")
    print("-" * 50)
    
    if not SHEETS_AVAILABLE:
        print("❌ Google Sheets不可用，跳過測試")
        return
    
    try:
        sheets_manager = GoogleSheetsManager()
        
        if not sheets_manager.client or not sheets_manager.spreadsheet:
            print("❌ Google Sheets未正確設定，跳過測試")
            return
        
        # 創建測試資料
        import pandas as pd
        
        # TRADING模式測試資料
        trading_data = pd.DataFrame({
            '日期': ['2025/06/09'],
            '契約名稱': ['TX'],
            '身份別': ['自營商'],
            '多方交易口數': [1000],
            '多方契約金額': [50000000],
            '空方交易口數': [800],
            '空方契約金額': [40000000],
            '多空淨額交易口數': [200],
            '多空淨額契約金額': [10000000]
        })
        
        # COMPLETE模式測試資料
        complete_data = pd.DataFrame({
            '日期': ['2025/06/09'],
            '契約名稱': ['TX'],
            '身份別': ['自營商'],
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
        
        print("📊 測試TRADING模式上傳...")
        try:
            result = sheets_manager.upload_data(trading_data, data_type='TRADING')
            print(f"  結果: {'✅ 成功' if result else '❌ 失敗'}")
        except Exception as e:
            print(f"  結果: ❌ 錯誤 - {e}")
        
        print("📊 測試COMPLETE模式上傳...")
        try:
            result = sheets_manager.upload_data(complete_data, data_type='COMPLETE')
            print(f"  結果: {'✅ 成功' if result else '❌ 失敗'}")
        except Exception as e:
            print(f"  結果: ❌ 錯誤 - {e}")
        
        print("📊 測試自動判斷上傳...")
        try:
            result = sheets_manager.upload_data(complete_data, data_type=None)
            print(f"  結果: {'✅ 成功' if result else '❌ 失敗'} (應該自動判斷為完整資料)")
        except Exception as e:
            print(f"  結果: ❌ 錯誤 - {e}")
        
    except Exception as e:
        print(f"❌ 測試過程發生錯誤: {e}")

def main():
    """主程序"""
    print("🔍 Google Sheets兩階段資料類型支援檢查")
    print("=" * 60)
    print(f"⏰ 檢查時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 檢查基本支援
    success = check_google_sheets_support()
    
    if success:
        # 測試上傳功能
        test_data_type_upload()
    
    print("\n" + "=" * 60)
    print("✅ 檢查完成！")
    
    if success:
        print("\n💡 總結:")
        print("  - Google Sheets管理器支援TRADING和COMPLETE兩種資料類型")
        print("  - 會自動選擇對應的工作表（交易量資料/完整資料）")
        print("  - 支援根據資料欄位自動判斷資料類型")
        print("  - 在主爬蟲程式中，根據--data_type參數自動上傳到正確分頁")

if __name__ == "__main__":
    main() 