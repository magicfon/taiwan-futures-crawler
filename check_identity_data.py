#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
檢查Google Sheets中的身份別資料結構
"""

from google_sheets_manager import GoogleSheetsManager
import pandas as pd

def check_identity_structure():
    """檢查身份別資料結構"""
    try:
        print('🔍 檢查Google Sheets中的身份別資料結構...')
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        ws = spreadsheet.worksheet('歷史資料')
        
        # 讀取前50行來檢查資料結構
        values = ws.get_all_values()
        headers = values[0]
        sample_data = values[1:51]  # 前50行樣本
        
        print(f'📋 欄位名稱: {headers}')
        print(f'📊 樣本資料 (前50行):')
        
        df = pd.DataFrame(sample_data, columns=headers)
        
        # 檢查身份別欄位
        if '身份別' in df.columns:
            identity_types = df['身份別'].unique()
            print(f'🏷️ 身份別類型: {list(identity_types)}')
            
            # 統計每種身份別的資料筆數
            for identity in identity_types:
                count = len(df[df['身份別'] == identity])
                print(f'   • {identity}: {count} 筆')
        
        # 檢查契約名稱
        if '契約名稱' in df.columns:
            contracts = df['契約名稱'].unique()
            print(f'📈 契約名稱: {list(contracts)}')
        
        # 檢查日期
        if '日期' in df.columns:
            dates = df['日期'].unique()
            print(f'📅 日期樣本 (前10個): {list(dates[:10])}')
        
        # 顯示幾行完整資料作為參考
        print(f'\n📋 前5行完整資料:')
        for i, row in df.head(5).iterrows():
            row_dict = dict(row)
            print(f'第{i+1}行:')
            for key, value in row_dict.items():
                print(f'  {key}: {value}')
            print()
        
        # 檢查是否同一天同一契約有多種身份別
        if '日期' in df.columns and '契約名稱' in df.columns and '身份別' in df.columns:
            print('\n🔍 檢查同一天同一契約的身份別分布:')
            
            # 取一個日期和契約作為例子
            sample_date = df['日期'].iloc[0]
            sample_contract = df['契約名稱'].iloc[0]
            
            filtered = df[(df['日期'] == sample_date) & (df['契約名稱'] == sample_contract)]
            
            print(f'📅 日期: {sample_date}, 契約: {sample_contract}')
            print(f'📊 該日該契約的身份別資料:')
            
            for _, row in filtered.iterrows():
                identity = row['身份別']
                # 找到數值欄位來顯示
                numeric_cols = ['多空淨額交易口數', '多空淨額未平倉口數', '交易口數', '未平倉口數']
                values_str = []
                for col in numeric_cols:
                    if col in row and str(row[col]).strip():
                        values_str.append(f'{col}:{row[col]}')
                
                print(f'  • {identity}: {", ".join(values_str)}')
        
        return True
        
    except Exception as e:
        print(f'❌ 錯誤: {e}')
        return False

if __name__ == "__main__":
    check_identity_structure() 