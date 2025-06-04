#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試歷史資料上傳功能
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd

def test_history_upload():
    print("🧪 測試歷史資料上傳功能...")
    
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
    
    # 連接Google Sheets
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        return
    
    # 連接到現有試算表
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1Ltv8zsQcCQ5SiaYKsgCDetNC-SqEMZP4V33S2nKMuWI"
    spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
    
    if not spreadsheet:
        print("❌ 連接試算表失敗")
        return
    
    print(f"✅ 成功連接試算表: {spreadsheet.title}")
    
    # 檢查現有工作表
    print("\n📋 現有工作表:")
    worksheet_names = []
    for worksheet in spreadsheet.worksheets():
        worksheet_names.append(worksheet.title)
        print(f"  - {worksheet.title}")
    
    # 檢查是否有歷史資料工作表
    if "歷史資料" not in worksheet_names:
        print("\n⚠️ 找不到「歷史資料」工作表，正在建立...")
        try:
            # 建立歷史資料工作表
            worksheet = spreadsheet.add_worksheet(title="歷史資料", rows=1000, cols=20)
            
            # 設定標題行
            headers = manager.get_history_headers()
            worksheet.update('A1', [headers])
            
            # 格式化標題行
            worksheet.format('A1:Z1', {
                'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.9},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
            })
            
            print("✅ 歷史資料工作表建立成功!")
            
        except Exception as e:
            print(f"❌ 建立歷史資料工作表失敗: {e}")
            return
    
    # 上傳資料到歷史資料工作表
    print("\n🚀 開始上傳到歷史資料工作表...")
    result = manager.upload_data(df)
    
    if result:
        print("✅ 歷史資料上傳成功!")
        print(f"🌐 查看結果: {manager.get_spreadsheet_url()}")
        
        # 測試取得最近30天資料
        print("\n📊 測試取得最近30天資料...")
        recent_data = manager.get_recent_data_for_report(30)
        if not recent_data.empty:
            print(f"✅ 成功取得最近30天資料: {len(recent_data)} 筆")
        else:
            print("⚠️ 最近30天資料為空，可能是日期格式問題")
    else:
        print("❌ 上傳失敗")

if __name__ == "__main__":
    test_history_upload() 