#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
測試修正後的圖表資料載入功能
"""

from chart_generator import ChartGenerator

def test_chart_data_loading():
    """測試圖表資料載入功能"""
    print("🔍 測試圖表資料載入功能...")
    
    chart_generator = ChartGenerator()
    df = chart_generator.load_data_from_google_sheets(30)
    
    print(f'📊 載入資料數量: {len(df)} 筆')
    
    if not df.empty:
        print(f'📅 日期範圍: {df["日期"].min()} 到 {df["日期"].max()}')
        print(f'📈 契約種類: {list(df["契約名稱"].unique())}')
        print(f'📊 身份別: {list(df["身份別"].unique()) if "身份別" in df.columns else "無身份別欄位"}')
        
        print('\n📋 最新資料樣本:')
        print(df.tail(10).to_string())
        
        # 檢查是否有足夠的資料生成圖表
        unique_dates = df['日期'].nunique()
        unique_contracts = df['契約名稱'].nunique()
        
        print(f'\n📊 圖表生成能力檢查:')
        print(f'   - 不同日期數: {unique_dates}')
        print(f'   - 不同契約數: {unique_contracts}')
        
        if unique_dates >= 7 and unique_contracts >= 1:
            print('   ✅ 資料充足，可以生成圖表')
        else:
            print('   ⚠️ 資料不足，建議至少7天資料且包含1個以上契約')
    else:
        print('❌ 沒有載入到任何資料')
        print('原因可能是:')
        print('  1. Google Sheets連線問題')
        print('  2. 工作表中沒有足夠的近期資料')
        print('  3. 資料格式不符合預期')

if __name__ == "__main__":
    test_chart_data_loading() 