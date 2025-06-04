#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試連接新的Google Sheets
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd

def test_new_sheet():
    """測試連接新的Google Sheets"""
    try:
        print('🔗 測試連接新的Google Sheets...')
        
        # 新的Google Sheets URL
        new_sheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
        
        gm = GoogleSheetsManager()
        
        # 連接到新的試算表
        spreadsheet = gm.connect_spreadsheet(new_sheet_url)
        
        if not spreadsheet:
            print('❌ 無法連接到新的Google Sheets')
            return False
        
        print(f'✅ 成功連接到: {spreadsheet.title}')
        print(f'🔗 網址: {new_sheet_url}')
        
        # 檢查所有工作表
        worksheets = spreadsheet.worksheets()
        print(f'\n📋 工作表列表:')
        for i, ws in enumerate(worksheets, 1):
            print(f'  {i}. {ws.title}')
        
        # 尋找包含資料的工作表
        for ws in worksheets:
            try:
                print(f'\n📖 檢查工作表: {ws.title}')
                all_values = ws.get_all_values()
                
                if len(all_values) < 2:
                    print(f'  ⚠️ 沒有足夠資料')
                    continue
                
                headers = all_values[0]
                print(f'  📋 欄位: {headers[:5]}...')  # 只顯示前5個欄位
                
                # 檢查是否有日期欄位
                if '日期' not in headers:
                    print(f'  ⚠️ 沒有日期欄位')
                    continue
                
                # 檢查資料筆數
                valid_rows = len([row for row in all_values[1:] if row[0].strip()])
                print(f'  📊 有效資料行數: {valid_rows}')
                
                if valid_rows == 0:
                    continue
                
                # 檢查2025/6/3資料
                june3_rows = []
                for i, row in enumerate(all_values[1:], 1):
                    if len(row) > 0 and '2025/6/3' in row[0]:
                        june3_rows.append((i, row))
                
                if june3_rows:
                    print(f'  ✅ 找到2025/6/3資料: {len(june3_rows)}筆')
                    for row_num, row in june3_rows[:3]:
                        print(f'    行{row_num}: {row[:3]}')
                else:
                    print(f'  ❌ 沒有2025/6/3資料')
                
                # 檢查最新日期
                dates = [row[0] for row in all_values[1:] if row[0].strip()]
                if dates:
                    print(f'  📅 第一個日期: {dates[0]}')
                    print(f'  📅 最後一個日期: {dates[-1]}')
                    
                    # 檢查最新5個日期
                    unique_dates = list(dict.fromkeys(dates))  # 去重保持順序
                    print(f'  📋 最新5個日期: {unique_dates[-5:]}')
                
            except Exception as e:
                print(f'  ❌ 檢查失敗: {e}')
                continue
        
        return True
        
    except Exception as e:
        print(f'❌ 測試失敗: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_new_sheet() 