#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
檢查Google Sheets中的資料範圍和最新日期
"""

from google_sheets_manager import GoogleSheetsManager
import json
from pathlib import Path
import pandas as pd

def check_sheets_data():
    """檢查Google Sheets中的資料情況"""
    try:
        # 載入試算表配置
        config_file = Path('config/spreadsheet_config.json')
        if not config_file.exists():
            print('❌ 找不到試算表配置檔案')
            return
            
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print(f'📊 試算表網址: {config.get("spreadsheet_url", "N/A")}')
        
        # 連接試算表
        sheets_manager = GoogleSheetsManager()
        sheets_manager.connect_spreadsheet(config.get('spreadsheet_id'))
        
        if not sheets_manager.spreadsheet:
            print('❌ 無法連接到試算表')
            return
        
        # 列出所有工作表及其資料情況
        worksheets = sheets_manager.spreadsheet.worksheets()
        print(f'\n📋 發現 {len(worksheets)} 個工作表:')
        
        for ws in worksheets:
            try:
                data = ws.get_all_records()
                if not data:
                    print(f'  - 📄 {ws.title}: 無資料')
                    continue
                
                df = pd.DataFrame(data)
                row_count = len(df)
                
                # 檢查日期欄位
                date_info = ""
                if '日期' in df.columns:
                    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                    valid_dates = df['日期'].dropna()
                    if not valid_dates.empty:
                        min_date = valid_dates.min().strftime('%Y-%m-%d')
                        max_date = valid_dates.max().strftime('%Y-%m-%d')
                        date_info = f" (日期範圍: {min_date} ~ {max_date})"
                    else:
                        date_info = " (無有效日期)"
                
                print(f'  - 📊 {ws.title}: {row_count} 筆資料{date_info}')
                
                # 如果是主要資料工作表，顯示更多詳情
                if ws.title in ['原始資料', '完整資料', '交易量資料'] and '日期' in df.columns:
                    # 統計各契約的資料數量
                    if '契約名稱' in df.columns:
                        contract_stats = df['契約名稱'].value_counts()
                        print(f'    契約統計: {dict(contract_stats)}')
                    
                    # 顯示最近5天的資料
                    recent_data = df.tail(10)
                    if not recent_data.empty:
                        print(f'    最近資料樣本:')
                        for _, row in recent_data.iterrows():
                            date_str = row.get('日期', 'N/A')
                            if pd.notna(date_str):
                                try:
                                    date_str = pd.to_datetime(date_str).strftime('%Y-%m-%d')
                                except:
                                    pass
                            contract = row.get('契約名稱', 'N/A')
                            identity = row.get('身份別', 'N/A')
                            print(f'      {date_str} | {contract} | {identity}')
                
            except Exception as e:
                print(f'  - ❌ {ws.title}: 讀取失敗 ({e})')
        
        # 特別檢查「原始資料」工作表
        print('\n🔍 詳細檢查「原始資料」工作表:')
        try:
            worksheet = sheets_manager.spreadsheet.worksheet("原始資料")
            data = worksheet.get_all_records()
            
            if data:
                df = pd.DataFrame(data)
                df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                
                print(f'   總資料量: {len(df)} 筆')
                print(f'   有效日期: {df["日期"].notna().sum()} 筆')
                
                valid_df = df[df['日期'].notna()]
                if not valid_df.empty:
                    print(f'   最早日期: {valid_df["日期"].min().strftime("%Y-%m-%d")}')
                    print(f'   最新日期: {valid_df["日期"].max().strftime("%Y-%m-%d")}')
                    
                    # 檢查最近7天的資料
                    from datetime import datetime, timedelta
                    today = datetime.now()
                    recent_7days = today - timedelta(days=7)
                    
                    recent_data = valid_df[valid_df['日期'] >= recent_7days]
                    print(f'   近7天資料: {len(recent_data)} 筆')
                    
                    if not recent_data.empty:
                        print('   近期日期明細:')
                        recent_dates = recent_data['日期'].dt.date.unique()
                        for date in sorted(recent_dates):
                            count = len(recent_data[recent_data['日期'].dt.date == date])
                            weekday = pd.to_datetime(date).strftime('%A')
                            print(f'     {date} ({weekday}): {count} 筆')
            else:
                print('   ❌ 「原始資料」工作表無資料')
                
        except Exception as e:
            print(f'   ❌ 無法讀取「原始資料」工作表: {e}')
    
    except Exception as e:
        print(f'❌ 檢查過程發生錯誤: {e}')

if __name__ == "__main__":
    check_sheets_data() 