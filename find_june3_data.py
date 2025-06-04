#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
專門搜尋6/3資料的腳本
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from datetime import datetime

def find_june3_data():
    """搜尋所有工作表中的6/3資料"""
    try:
        print('🔍 搜尋所有工作表中的6/3資料...')
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 檢查所有工作表
        worksheets = spreadsheet.worksheets()
        print(f"📊 將檢查以下工作表: {[ws.title for ws in worksheets]}")
        
        june3_found = False
        
        for ws in worksheets:
            try:
                print(f"\n📖 檢查工作表: {ws.title}")
                all_values = ws.get_all_values()
                
                if len(all_values) < 2:
                    print(f"  ⚠️ {ws.title} 沒有足夠資料（只有{len(all_values)}行）")
                    continue
                
                headers = all_values[0]
                print(f"  📋 欄位: {headers}")
                
                # 轉換為DataFrame
                df = pd.DataFrame(all_values[1:], columns=headers)
                df = df[df.iloc[:, 0].astype(str).str.strip() != '']  # 過濾空行，使用第一欄
                
                print(f"  📊 有效資料行數: {len(df)}")
                
                if len(df) == 0:
                    print(f"  ⚠️ {ws.title} 沒有有效資料")
                    continue
                
                # 檢查是否有日期欄位
                date_columns = ['日期', 'date', '交易日期', 'Date']
                date_col = None
                
                for col in date_columns:
                    if col in headers:
                        date_col = col
                        break
                
                if not date_col:
                    print(f"  ⚠️ {ws.title} 沒有找到日期欄位")
                    continue
                
                print(f"  📅 使用日期欄位: {date_col}")
                
                # 嘗試轉換日期並搜尋6/3
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce', 
                                            format='%Y/%m/%d', infer_datetime_format=True)
                
                # 過濾有效日期
                valid_dates = df[date_col].notna()
                df_valid = df[valid_dates]
                
                if len(df_valid) == 0:
                    print(f"  ⚠️ {ws.title} 沒有有效的日期資料")
                    continue
                
                print(f"  📅 日期範圍: {df_valid[date_col].min()} 到 {df_valid[date_col].max()}")
                
                # 搜尋6/3資料（嘗試2024和2025年）
                june3_dates = [
                    pd.to_datetime('2025/06/03').date(),
                    pd.to_datetime('2024/06/03').date()
                ]
                
                for target_date in june3_dates:
                    june3_data = df_valid[df_valid[date_col].dt.date == target_date]
                    
                    if not june3_data.empty:
                        print(f"  ✅ 在 {ws.title} 找到 {target_date} 的資料: {len(june3_data)}筆")
                        june3_found = True
                        
                        # 顯示找到的資料詳細內容
                        print(f"  📋 {target_date} 資料內容:")
                        for i, (_, row) in enumerate(june3_data.head(10).iterrows()):
                            row_info = []
                            for col in headers[:5]:  # 只顯示前5個欄位
                                if col in row and str(row[col]).strip():
                                    row_info.append(f"{col}:{row[col]}")
                            print(f"    {i+1}. {' | '.join(row_info)}")
                            
                        if len(june3_data) > 10:
                            print(f"    ... 還有 {len(june3_data)-10} 筆資料")
                
                # 顯示最新幾個日期
                latest_dates = sorted(df_valid[date_col].dt.date.unique(), reverse=True)
                print(f"  📆 最新5個日期: {latest_dates[:5]}")
                
            except Exception as e:
                print(f"  ❌ {ws.title} 讀取失敗: {e}")
                continue
        
        if not june3_found:
            print("\n❌ 在所有工作表中都沒有找到6/3資料")
            print("可能的原因:")
            print("1. 6/3是非交易日")
            print("2. 資料尚未更新")
            print("3. 日期格式不同")
        else:
            print(f"\n✅ 已找到6/3資料！")
        
        return june3_found
        
    except Exception as e:
        print(f'❌ 搜尋失敗: {e}')
        return False

if __name__ == "__main__":
    find_june3_data() 