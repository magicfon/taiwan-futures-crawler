#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
檢查Google Sheets的完整結構
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from collections import Counter

def check_structure():
    print('🔍 檢查Google Sheets完整結構...')
    
    try:
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        ws = spreadsheet.worksheet('歷史資料')
        
        print('📊 讀取完整工作表...')
        all_values = ws.get_all_values()
        
        print(f'總行數: {len(all_values)}')
        print(f'標題行: {all_values[0]}')
        
        # 統計所有日期
        dates = []
        for i in range(1, len(all_values)):
            if len(all_values[i]) > 0 and all_values[i][0].strip():
                dates.append(all_values[i][0])
        
        print(f'\n📅 有效日期行數: {len(dates)}')
        
        # 按年份統計
        year_counts = Counter()
        month_counts = Counter()
        
        for date_str in dates:
            try:
                if '/' in date_str:
                    parts = date_str.split('/')
                    if len(parts) >= 3:
                        year = parts[0]
                        month = parts[1]
                        year_counts[year] += 1
                        month_counts[f'{year}/{month}'] += 1
            except:
                continue
        
        print(f'\n📈 按年份統計:')
        for year in sorted(year_counts.keys()):
            print(f'  {year}: {year_counts[year]}筆')
        
        print(f'\n📈 2025年按月份統計:')
        for month_key in sorted([k for k in month_counts.keys() if k.startswith('2025')]):
            print(f'  {month_key}: {month_counts[month_key]}筆')
        
        # 檢查具體的6月資料
        june_2025_dates = [d for d in dates if d.startswith('2025/6/') or d.startswith('2025/06/')]
        print(f'\n📅 2025年6月的所有日期:')
        if june_2025_dates:
            unique_june_dates = sorted(set(june_2025_dates))
            for date in unique_june_dates:
                count = june_2025_dates.count(date)
                print(f'  {date}: {count}筆')
        else:
            print('  ❌ 沒有2025年6月的資料')
        
        # 檢查最新和最舊的日期
        print(f'\n📆 日期範圍分析:')
        print(f'第一個日期: {dates[0] if dates else "無"}')
        print(f'最後一個日期: {dates[-1] if dates else "無"}')
        
        # 顯示最新10個唯一日期
        unique_dates = list(dict.fromkeys(dates))  # 保持順序去重
        print(f'\n📋 最新10個唯一日期:')
        for i, date in enumerate(unique_dates[-10:], 1):
            print(f'  {i:2d}. {date}')
        
        # 搜尋所有6/3
        all_june3 = [d for d in dates if '/6/3' in d or '/06/03' in d]
        print(f'\n🔍 所有6/3相關日期:')
        if all_june3:
            for date in set(all_june3):
                count = all_june3.count(date)
                print(f'  {date}: {count}筆')
        else:
            print('  ❌ 沒有任何6/3的資料')
            
        return True
        
    except Exception as e:
        print(f'❌ 錯誤: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    check_structure() 