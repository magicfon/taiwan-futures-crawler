#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試Google Sheets上傳功能
"""

import pandas as pd
from google_sheets_manager import GoogleSheetsManager
from pathlib import Path

def main():
    print("🧪 測試Google Sheets上傳功能...")
    
    # 建立Google Sheets管理器
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        return
    
    # 連接到我們的試算表
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1Ltv8zsQcCQ5SiaYKsgCDetNC-SqEMZP4V33S2nKMuWI"
    spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
    
    if not spreadsheet:
        print("❌ 連接試算表失敗")
        return
    
    print(f"✅ 成功連接到試算表: {spreadsheet.title}")
    
    # 查看現有的CSV檔案
    csv_files = list(Path("output").glob("*.csv"))
    if not csv_files:
        print("❌ 找不到CSV檔案來測試")
        return
    
    # 選擇最新的CSV檔案
    latest_csv = max(csv_files, key=lambda x: x.stat().st_mtime)
    print(f"📁 使用測試檔案: {latest_csv}")
    
    # 讀取CSV資料
    try:
        df = pd.read_csv(latest_csv, encoding='utf-8')
        print(f"📊 讀取到 {len(df)} 筆資料")
        print(f"📋 欄位: {list(df.columns)}")
        
        # 顯示前幾筆資料
        print(f"\n🔍 前3筆資料預覽:")
        print(df.head(3).to_string())
        
    except Exception as e:
        print(f"❌ 讀取CSV失敗: {e}")
        return
    
    # 測試上傳到Google Sheets
    print(f"\n🚀 開始上傳到Google Sheets...")
    
    try:
        # 上傳主要資料
        if manager.upload_data(df, "最新30天資料"):
            print("✅ 主要資料上傳成功")
        else:
            print("❌ 主要資料上傳失敗")
        
        # 更新系統資訊
        if manager.update_system_info():
            print("✅ 系統資訊更新成功")
        else:
            print("❌ 系統資訊更新失敗")
        
        print(f"\n🌐 試算表網址: {manager.get_spreadsheet_url()}")
        print(f"🎉 測試完成！請打開試算表查看資料")
        
    except Exception as e:
        print(f"❌ 上傳過程發生錯誤: {e}")

if __name__ == "__main__":
    main() 