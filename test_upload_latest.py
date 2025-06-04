#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試最新資料上傳到Google Sheets
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd

def test_upload():
    print("🧪 測試最新資料上傳到Google Sheets...")
    
    # 讀取最新的CSV檔案
    csv_file = "output/taifex_20241231-20241231_TX_自營商_投信_外資.csv"
    df = pd.read_csv(csv_file)
    
    print(f"📊 讀取檔案: {csv_file}")
    print(f"📊 資料筆數: {len(df)}")
    print(f"📊 欄位: {list(df.columns)}")
    
    # 顯示資料預覽
    print("\n📋 資料預覽:")
    for _, row in df.iterrows():
        print(f"  {row['身份別']}: 多方{row['多方交易口數']}口 空方{row['空方交易口數']}口")
    
    # 上傳到Google Sheets
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        return
    
    print("\n🚀 開始上傳到Google Sheets...")
    result = manager.upload_data(df)
    
    if result:
        print("✅ Google Sheets上傳成功!")
        print(f"🌐 查看結果: {manager.get_spreadsheet_url()}")
    else:
        print("❌ 上傳失敗")

if __name__ == "__main__":
    test_upload() 