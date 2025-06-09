#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
將CSV檔案的資料直接上傳到Google Sheets歷史資料工作表
"""

from google_sheets_manager import GoogleSheetsManager
import json
from pathlib import Path
import pandas as pd
import sys

def main(csv_file_path=None):
    """
    上傳CSV檔案到Google Sheets
    
    Args:
        csv_file_path (str): CSV檔案路徑，如果沒有指定則使用最新的2024年12月資料
    """
    
    print("🚀 開始上傳CSV資料到Google Sheets歷史資料工作表...")
    
    # 如果沒有指定檔案，使用最新的2024年12月資料
    if not csv_file_path:
        csv_file_path = "output/taifex_20241220-20241231_TX_TE_MTX_ZMX_NQF_自營商_投信_外資.csv"
    
    csv_file = Path(csv_file_path)
    if not csv_file.exists():
        print(f"❌ 檔案不存在: {csv_file_path}")
        return False
    
    print(f"📂 讀取檔案: {csv_file_path}")
    
    # 讀取CSV檔案
    try:
        df = pd.read_csv(csv_file_path, encoding='utf-8')
        print(f"📊 成功讀取 {len(df)} 筆資料")
    except Exception as e:
        print(f"❌ 讀取CSV檔案失敗: {e}")
        return False
    
    # 初始化Google Sheets管理器
    sheets_manager = GoogleSheetsManager()
    
    if not sheets_manager.client:
        print("❌ Google Sheets認證失敗")
        return False
    
    # 連接到試算表
    config_file = Path('config/spreadsheet_config.json')
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except Exception as e:
        print(f"❌ 讀取配置檔案失敗: {e}")
        return False
    
    sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
    if not sheets_manager.spreadsheet:
        print("❌ 無法連接到試算表")
        return False
    
    print(f"✅ 已連接到試算表: {sheets_manager.spreadsheet.title}")
    
    # 顯示將要上傳的資料概覽
    dates = df['日期'].unique()
    print(f"📅 資料涵蓋日期: {len(dates)} 天")
    print(f"📅 日期範圍: {min(dates)} ~ {max(dates)}")
    
    # 按日期統計
    for date in sorted(dates):
        date_data = df[df['日期'] == date]
        contracts = date_data['契約名稱'].unique()
        identities = date_data['身份別'].unique()
        print(f"   📅 {date}: {len(date_data)} 筆 (契約: {len(contracts)}, 身份: {len(identities)})")
    
    # 上傳到Google Sheets
    print(f"\n🚀 開始上傳到歷史資料工作表...")
    try:
        success = sheets_manager.upload_data(df, worksheet_name="歷史資料")
        if success:
            print("✅ 資料上傳成功！")
            print(f"🌐 試算表網址: {sheets_manager.get_spreadsheet_url()}")
            
            # 檢查上傳結果
            print(f"\n🔍 檢查上傳結果...")
            worksheet = sheets_manager.spreadsheet.worksheet('歷史資料')
            all_data = worksheet.get_all_values()
            
            # 統計上傳後的資料
            uploaded_dates = set()
            for row in all_data[1:]:  # 跳過標題行
                if len(row) > 0 and row[0]:
                    date_str = row[0]
                    uploaded_dates.add(date_str)
            
            print(f"   📊 試算表中總共有 {len(all_data)-1} 筆資料")
            print(f"   📅 涵蓋 {len(uploaded_dates)} 個不同日期")
            
            # 檢查是否包含我們剛上傳的日期
            found_dates = 0
            for date in dates:
                if date in uploaded_dates:
                    found_dates += 1
            
            print(f"   ✅ 成功確認 {found_dates}/{len(dates)} 個日期的資料已上傳")
            
            return True
        else:
            print("❌ 資料上傳失敗")
            return False
    except Exception as e:
        print(f"❌ 上傳過程發生錯誤: {e}")
        return False

if __name__ == "__main__":
    # 支援命令行參數
    csv_file_path = None
    if len(sys.argv) > 1:
        csv_file_path = sys.argv[1]
    
    success = main(csv_file_path)
    sys.exit(0 if success else 1) 