#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
詳細檢查2025年6/3資料的腳本
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from datetime import datetime, timedelta
import calendar

def check_june3_detailed():
    """詳細檢查2025年6/3的情況"""
    try:
        print('📅 詳細檢查2025年6/3資料...')
        
        # 檢查2025年6/3是星期幾
        june3_2025 = datetime(2025, 6, 3)
        weekday = june3_2025.weekday()  # 0=週一, 6=週日
        weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
        
        print(f"📆 2025年6月3日是 {weekday_names[weekday]}")
        
        if weekday >= 5:  # 週六或週日
            print(f"⚠️ 2025年6月3日是{weekday_names[weekday]}，為非交易日")
        else:
            print(f"✅ 2025年6月3日是{weekday_names[weekday]}，應為交易日")
        
        # 檢查Google Sheets資料
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        ws = spreadsheet.worksheet('歷史資料')
        
        all_values = ws.get_all_values()
        headers = all_values[0]
        df = pd.DataFrame(all_values[1:], columns=headers)
        df = df[df['日期'].str.strip() != '']
        
        # 轉換日期
        df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d', errors='coerce')
        df = df.dropna(subset=['日期'])
        
        print(f"\n📊 Google Sheets 資料概況:")
        print(f"總資料筆數: {len(df)}")
        print(f"日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
        
        # 檢查6月前後的資料
        june_2025_data = df[
            (df['日期'].dt.year == 2025) & 
            (df['日期'].dt.month == 6)
        ]
        
        print(f"\n📅 2025年6月的資料:")
        if june_2025_data.empty:
            print("❌ 沒有2025年6月的資料")
        else:
            june_dates = sorted(june_2025_data['日期'].dt.date.unique())
            print(f"✅ 有2025年6月資料，日期包含: {june_dates}")
            
            # 特別檢查6/3
            june3_data = june_2025_data[june_2025_data['日期'].dt.date == june3_2025.date()]
            if june3_data.empty:
                print("❌ 沒有2025年6月3日的資料")
            else:
                print(f"✅ 有2025年6月3日資料: {len(june3_data)}筆")
        
        # 檢查最新幾天的資料
        print(f"\n📆 最新10個交易日:")
        latest_dates = sorted(df['日期'].dt.date.unique(), reverse=True)[:10]
        for i, date in enumerate(latest_dates, 1):
            weekday_name = weekday_names[date.weekday()]
            print(f"{i:2d}. {date} ({weekday_name})")
        
        # 檢查是否有比6/2更新的資料
        after_june2 = df[df['日期'].dt.date > datetime(2025, 6, 2).date()]
        if not after_june2.empty:
            print(f"\n📈 比2025/6/2更新的資料:")
            new_dates = sorted(after_june2['日期'].dt.date.unique())
            for date in new_dates:
                weekday_name = weekday_names[date.weekday()]
                count = len(after_june2[after_june2['日期'].dt.date == date])
                print(f"  {date} ({weekday_name}): {count}筆")
        else:
            print(f"\n📊 沒有比2025/6/2更新的資料")
        
        # 檢查原始文字格式
        print(f"\n🔍 檢查原始日期格式...")
        raw_dates = [row[0] for row in all_values[1:] if row[0].strip()]
        june3_raw = [date for date in raw_dates if '2025/6/3' in date or '2025/06/03' in date]
        
        if june3_raw:
            print(f"✅ 在原始資料中找到2025/6/3格式: {june3_raw[:5]}")
        else:
            print(f"❌ 在原始資料中沒有找到2025/6/3格式")
            
        # 搜尋類似的日期
        similar_dates = [date for date in raw_dates if '6/3' in date or '/06/03' in date]
        if similar_dates:
            print(f"📋 包含6/3的所有日期: {similar_dates[:10]}")
        
        return True
        
    except Exception as e:
        print(f'❌ 檢查失敗: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    check_june3_detailed() 