#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
簡單檢查2025年6/3資料
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd

def simple_check():
    print('🔍 簡單檢查2025年6/3資料...')
    
    try:
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        ws = spreadsheet.worksheet('歷史資料')
        
        print('📊 讀取工作表資料...')
        all_values = ws.get_all_values()
        print(f'總行數: {len(all_values)}')
        
        # 檢查原始資料中是否有2025/6/3
        june3_rows = []
        for i, row in enumerate(all_values):
            if len(row) > 0 and '2025/6/3' in row[0]:
                june3_rows.append((i, row))
        
        if june3_rows:
            print(f'✅ 找到2025/6/3原始資料：{len(june3_rows)}行')
            for row_num, row in june3_rows[:5]:
                print(f'  行{row_num}: {row[:3]}')
        else:
            print('❌ 原始資料中沒有2025/6/3')
            
        # 檢查包含6/3的所有行
        all_june3 = []
        for i, row in enumerate(all_values):
            if len(row) > 0 and '6/3' in row[0]:
                all_june3.append((i, row[0]))
        
        print(f'\n📋 所有包含6/3的日期：')
        for row_num, date in all_june3:
            print(f'  行{row_num}: {date}')
            
        # 檢查最新日期
        print(f'\n📅 最新10行的日期：')
        for i in range(1, min(11, len(all_values))):
            if len(all_values[i]) > 0:
                print(f'  行{i}: {all_values[i][0]}')
                
        return len(june3_rows) > 0
        
    except Exception as e:
        print(f'❌ 錯誤: {e}')
        return False

if __name__ == "__main__":
    simple_check() 