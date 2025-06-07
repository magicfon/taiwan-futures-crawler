#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
專門檢查2025年6/2~6/5日期範圍的資料
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from datetime import datetime, timedelta
import calendar

def check_june_2_to_5():
    """檢查2025年6/2~6/5的資料狀況"""
    try:
        print('📅 檢查2025年6/2~6/5日期範圍的資料...')
        
        # 先檢查這些日期是星期幾
        target_dates = []
        for day in range(2, 6):  # 6/2 到 6/5
            date = datetime(2025, 6, day)
            weekday = date.weekday()  # 0=週一, 6=週日
            weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
            is_trading_day = weekday < 5  # 週一到週五是交易日
            
            target_dates.append({
                'date': date,
                'date_str': f"2025/6/{day}",
                'weekday': weekday_names[weekday],
                'is_trading_day': is_trading_day
            })
            
            print(f"📆 2025年6月{day}日是 {weekday_names[weekday]} {'✅交易日' if is_trading_day else '❌非交易日'}")
        
        # 檢查Google Sheets資料
        print(f"\n🔍 檢查Google Sheets資料...")
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 檢查所有工作表
        worksheets = spreadsheet.worksheets()
        print(f"📊 可用工作表: {[ws.title for ws in worksheets]}")
        
        # 檢查歷史資料工作表
        ws = spreadsheet.worksheet('歷史資料')
        all_values = ws.get_all_values()
        
        print(f"📋 工作表總行數: {len(all_values)}")
        
        if len(all_values) < 2:
            print("❌ 工作表沒有足夠資料")
            return False
        
        headers = all_values[0]
        print(f"📋 欄位: {headers}")
        
        # 檢查原始資料中的日期格式
        print(f"\n🔍 檢查原始日期格式...")
        raw_dates = []
        for i, row in enumerate(all_values[1:], 1):
            if len(row) > 0 and row[0].strip():
                raw_dates.append((i, row[0]))
        
        print(f"📊 有效日期行數: {len(raw_dates)}")
        
        if len(raw_dates) > 0:
            print(f"📅 前10個日期: {[date for _, date in raw_dates[:10]]}")
            print(f"📅 後10個日期: {[date for _, date in raw_dates[-10:]]}")
        
        # 搜尋6/2~6/5的原始資料
        found_dates = {}
        for target in target_dates:
            date_str = target['date_str']
            found_rows = []
            
            for row_num, raw_date in raw_dates:
                if date_str in raw_date or f"2025/06/0{target['date'].day}" in raw_date:
                    found_rows.append((row_num, raw_date))
            
            found_dates[date_str] = found_rows
            
            if found_rows:
                print(f"✅ 找到 {date_str} ({target['weekday']}) 原始資料: {len(found_rows)}行")
                for row_num, raw_date in found_rows[:3]:
                    print(f"   行{row_num}: {raw_date}")
            else:
                print(f"❌ 沒有找到 {date_str} ({target['weekday']}) 原始資料")
        
        # 嘗試轉換為DataFrame並分析
        print(f"\n📊 嘗試轉換為DataFrame...")
        try:
            df = pd.DataFrame(all_values[1:], columns=headers)
            df = df[df['日期'].str.strip() != '']
            
            print(f"📊 過濾後有效行數: {len(df)}")
            
            if len(df) > 0:
                # 嘗試轉換日期
                df['日期_原始'] = df['日期']
                df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d', errors='coerce')
                
                valid_dates = df['日期'].notna()
                df_valid = df[valid_dates]
                
                print(f"📅 成功轉換日期的行數: {len(df_valid)}")
                
                if len(df_valid) > 0:
                    print(f"📅 日期範圍: {df_valid['日期'].min()} 到 {df_valid['日期'].max()}")
                    
                    # 檢查6/2~6/5的資料
                    for target in target_dates:
                        target_date = target['date'].date()
                        date_data = df_valid[df_valid['日期'].dt.date == target_date]
                        
                        if not date_data.empty:
                            print(f"✅ DataFrame中找到 {target['date_str']} 資料: {len(date_data)}筆")
                            
                            # 顯示資料詳細內容
                            for i, (_, row) in enumerate(date_data.head(5).iterrows()):
                                print(f"   {i+1}. 契約:{row.get('契約名稱', 'N/A')} 身份:{row.get('身份別', 'N/A')} 交易:{row.get('多空淨額交易口數', 'N/A')}")
                        else:
                            print(f"❌ DataFrame中沒有 {target['date_str']} 資料")
                
                # 檢查日期轉換失敗的資料
                invalid_dates = df[~valid_dates]
                if len(invalid_dates) > 0:
                    print(f"\n⚠️ 日期轉換失敗的資料: {len(invalid_dates)}筆")
                    print("前5個失敗的日期格式:")
                    for i, date_str in enumerate(invalid_dates['日期_原始'].head(5)):
                        print(f"   {i+1}. '{date_str}'")
            
        except Exception as e:
            print(f"❌ DataFrame轉換失敗: {e}")
        
        # 檢查其他工作表
        print(f"\n🔍 檢查其他工作表...")
        for ws_name in ['每日摘要', '三大法人趨勢']:
            try:
                if ws_name in [ws.title for ws in worksheets]:
                    ws_other = spreadsheet.worksheet(ws_name)
                    other_values = ws_other.get_all_values()
                    print(f"📊 {ws_name}: {len(other_values)}行")
                    
                    if len(other_values) > 1:
                        # 檢查是否有6/2~6/5的資料
                        for target in target_dates:
                            found_in_other = False
                            for row in other_values[1:]:
                                if len(row) > 0 and target['date_str'] in str(row[0]):
                                    found_in_other = True
                                    break
                            
                            if found_in_other:
                                print(f"   ✅ {ws_name}中有 {target['date_str']} 資料")
                            else:
                                print(f"   ❌ {ws_name}中沒有 {target['date_str']} 資料")
            except Exception as e:
                print(f"   ❌ 檢查{ws_name}失敗: {e}")
        
        return True
        
    except Exception as e:
        print(f'❌ 檢查失敗: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    check_june_2_to_5() 