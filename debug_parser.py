#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
詳細除錯解析器
"""

import requests
from bs4 import BeautifulSoup

def debug_parser():
    print("🔍 詳細除錯解析器...")
    
    BASE_URL = "https://www.taifex.com.tw/cht/3/futContractsDate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    # 測試2024/12/31的TX資料
    params = {
        'queryType': '2',
        'marketCode': '0',
        'dateaddcnt': '',
        'commodity_id': 'TX',
        'queryDate': '2024/12/31'
    }
    
    response = requests.get(BASE_URL, params=params, headers=headers, timeout=10)
    html_content = response.text
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 找到主要表格
    tables = soup.find_all('table')
    print(f"📊 找到 {len(tables)} 個表格")
    
    main_table = max(tables, key=lambda t: len(str(t)))
    print(f"📊 主要表格有 {len(main_table.find_all('tr'))} 行")
    
    # 詳細檢查每一行
    print("\n🔍 詳細檢查表格內容:")
    for i, row in enumerate(main_table.find_all('tr')[:10]):  # 只檢查前10行
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
            
        print(f"\n行 {i+1}:")
        for j, cell in enumerate(cells[:8]):  # 只顯示前8列
            text = cell.get_text(strip=True)
            print(f"  列{j+1}: '{text}'")
        
        # 檢查是否包含TX相關資訊
        row_text = ' '.join([cell.get_text(strip=True) for cell in cells])
        
        if 'TX' in row_text or '臺股期貨' in row_text or '台股期貨' in row_text:
            print(f"  ⭐ 這行可能包含TX資料!")
            print(f"  完整行文本: {row_text}")
            
            # 檢查是否有數字
            import re
            numbers = re.findall(r'[\d,]+', row_text)
            if numbers:
                print(f"  找到數字: {numbers[:5]}")

if __name__ == "__main__":
    debug_parser() 