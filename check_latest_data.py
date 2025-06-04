#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
檢查Google Sheets最新資料
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from datetime import datetime

def check_latest_data():
    """檢查最新資料狀況"""
    try:
        print('🔍 檢查Google Sheets最新資料...')
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 檢查所有工作表
        worksheets = spreadsheet.worksheets()
        print(f"📊 可用工作表: {[ws.title for ws in worksheets]}")
        
        # 檢查歷史資料工作表
        ws = spreadsheet.worksheet('歷史資料')
        all_values = ws.get_all_values()
        
        if len(all_values) < 2:
            print("❌ 工作表沒有資料")
            return
        
        headers = all_values[0]
        print(f"📋 欄位: {headers}")
        
        # 轉換為DataFrame
        df = pd.DataFrame(all_values[1:], columns=headers)
        df = df[df['日期'].str.strip() != '']
        
        # 轉換日期
        df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d')
        
        # 檢查日期範圍
        print(f"📅 總資料筆數: {len(df)}")
        print(f"📅 日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
        
        # 檢查最新幾天的資料
        latest_dates = df['日期'].dt.date.unique()
        latest_dates_sorted = sorted(latest_dates, reverse=True)
        
        print(f"📆 最新5個日期:")
        for i, date in enumerate(latest_dates_sorted[:5]):
            date_data = df[df['日期'].dt.date == date]
            contracts = date_data['契約名稱'].unique()
            identities = date_data['身份別'].unique()
            print(f"  {i+1}. {date}: {len(date_data)}筆 契約:{list(contracts)} 身份:{list(identities)}")
        
        # 檢查6/3資料
        target_date = pd.to_datetime('2025/06/03').date()
        june_3_data = df[df['日期'].dt.date == target_date]
        
        if not june_3_data.empty:
            print(f"\n✅ 找到6/3資料: {len(june_3_data)}筆")
            print("6/3資料詳細:")
            for _, row in june_3_data.head(10).iterrows():
                print(f"  契約:{row['契約名稱']} 身份:{row['身份別']} 交易:{row['多空淨額交易口數']} 未平倉:{row['多空淨額未平倉口數']}")
        else:
            print(f"\n❌ 沒有找到6/3資料")
        
        # 檢查數值範圍是否合理
        print(f"\n🔢 數值檢查:")
        numeric_cols = ['多空淨額交易口數', '多空淨額未平倉口數']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
                print(f"{col}:")
                print(f"  範圍: {df[col].min()} 到 {df[col].max()}")
                print(f"  平均: {df[col].mean():.0f}")
        
        # 檢查TX契約最新資料
        print(f"\n💼 TX契約最新資料:")
        tx_data = df[df['契約名稱'] == 'TX'].sort_values('日期', ascending=False)
        if not tx_data.empty:
            latest_tx = tx_data.head(3)
            for _, row in latest_tx.iterrows():
                print(f"  {row['日期'].strftime('%Y-%m-%d')} {row['身份別']}: 交易{row['多空淨額交易口數']} 未平倉{row['多空淨額未平倉口數']}")
        
        return True
        
    except Exception as e:
        print(f'❌ 錯誤: {e}')
        return False

if __name__ == "__main__":
    check_latest_data() 