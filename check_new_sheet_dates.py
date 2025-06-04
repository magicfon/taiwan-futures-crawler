#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
詳細檢查新Google Sheets的日期格式
"""

from google_sheets_manager import GoogleSheetsManager

def check_new_sheet_dates():
    """詳細檢查新Google Sheets的日期"""
    try:
        print('🔍 詳細檢查新Google Sheets的日期格式...')
        
        new_sheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.connect_spreadsheet(new_sheet_url)
        
        ws = spreadsheet.worksheet('歷史資料')
        all_values = ws.get_all_values()
        
        print(f'總行數: {len(all_values)}')
        
        # 檢查最後10行的原始資料
        print(f'\n📋 最後10行的原始日期格式:')
        for i in range(max(1, len(all_values)-10), len(all_values)):
            if i < len(all_values) and len(all_values[i]) > 0:
                date_str = all_values[i][0]
                contract = all_values[i][1] if len(all_values[i]) > 1 else "無"
                identity = all_values[i][2] if len(all_values[i]) > 2 else "無"
                print(f'  行{i}: 日期="{date_str}" 契約={contract} 身份={identity}')
        
        # 搜尋包含"6/3"的所有行（不管年份）
        print(f'\n🔍 搜尋所有包含"6/3"的行:')
        june3_found = []
        for i, row in enumerate(all_values):
            if len(row) > 0:
                date_str = row[0]
                if '6/3' in date_str or '06/03' in date_str:
                    june3_found.append((i, date_str, row[1:3] if len(row) > 2 else []))
        
        if june3_found:
            print(f'✅ 找到{len(june3_found)}行包含6/3的資料:')
            for row_num, date_str, extra_info in june3_found:
                print(f'  行{row_num}: "{date_str}" {extra_info}')
        else:
            print('❌ 沒有找到包含6/3的資料')
        
        # 搜尋包含"2025"的最新日期
        print(f'\n📅 搜尋2025年的最新日期:')
        dates_2025 = []
        for row in all_values[1:]:  # 跳過標題行
            if len(row) > 0 and row[0].strip() and '2025' in row[0]:
                dates_2025.append(row[0])
        
        if dates_2025:
            unique_dates_2025 = list(dict.fromkeys(dates_2025))  # 去重保持順序
            print(f'  📊 2025年共有{len(unique_dates_2025)}個不同日期')
            print(f'  📋 最新10個2025年日期: {unique_dates_2025[-10:]}')
            
            # 特別檢查是否有2025/06/03或2025/6/3
            june3_2025_variants = [
                '2025/6/3', '2025/06/03', '2025-6-3', '2025-06-03'
            ]
            
            for variant in june3_2025_variants:
                if variant in unique_dates_2025:
                    print(f'  ✅ 找到日期格式: {variant}')
                elif variant in dates_2025:  # 檢查重複的
                    count = dates_2025.count(variant)
                    print(f'  ✅ 找到日期格式: {variant} (出現{count}次)')
        
        # 檢查原始資料的最後幾行是否真的是6/3
        print(f'\n🔎 檢查最後50行中的6月資料:')
        last_50_rows = all_values[-50:]
        june_rows = []
        for i, row in enumerate(last_50_rows):
            if len(row) > 0 and row[0].strip() and '/6/' in row[0]:
                june_rows.append((len(all_values)-50+i, row[0], row[1:3] if len(row) > 2 else []))
        
        if june_rows:
            print(f'  ✅ 最後50行中找到{len(june_rows)}行6月資料:')
            for row_num, date_str, extra_info in june_rows[-10:]:  # 顯示最後10個
                print(f'    行{row_num}: "{date_str}" {extra_info}')
        else:
            print(f'  ❌ 最後50行中沒有6月資料')
        
        return True
        
    except Exception as e:
        print(f'❌ 檢查失敗: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    check_new_sheet_dates() 