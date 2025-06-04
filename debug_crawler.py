#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
除錯台期所爬蟲，檢查網站回應內容
"""

import requests
from bs4 import BeautifulSoup
import datetime

def test_taifex_response():
    print("🔍 測試台期所網站回應...")
    
    # 設定請求參數
    BASE_URL = "https://www.taifex.com.tw/cht/3/futContractsDate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1'
    }
    
    # 測試不同的日期
    test_dates = [
        "2024/12/31",  # 最近的交易日
        "2024/12/30",  # 前一天
        "2024/12/27",  # 週五
        "2024/05/20",  # 你之前有資料的日期
    ]
    
    for date_str in test_dates:
        print(f"\n📅 測試日期: {date_str}")
        
        # 測試TX契約
        params = {
            'queryType': '2',
            'marketCode': '0',
            'dateaddcnt': '',
            'commodity_id': 'TX',
            'queryDate': date_str
        }
        
        try:
            response = requests.get(BASE_URL, params=params, headers=headers, timeout=10)
            print(f"   狀態碼: {response.status_code}")
            
            if response.status_code == 200:
                # 檢查回應內容
                html_content = response.text
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # 尋找表格
                tables = soup.find_all('table')
                print(f"   找到表格數量: {len(tables)}")
                
                # 檢查是否有無資料訊息
                no_data_messages = ['查無資料', '無交易資料', '尚無資料', 'No data']
                has_no_data = any(msg in html_content for msg in no_data_messages)
                print(f"   無資料訊息: {'是' if has_no_data else '否'}")
                
                # 顯示表格內容摘要
                if tables:
                    main_table = max(tables, key=lambda t: len(str(t)))
                    rows = main_table.find_all('tr')
                    print(f"   主要表格行數: {len(rows)}")
                    
                    # 顯示前幾行內容
                    for i, row in enumerate(rows[:5]):
                        cells = row.find_all(['td', 'th'])
                        if cells:
                            row_text = ' | '.join([cell.get_text(strip=True) for cell in cells[:5]])
                            print(f"   行{i+1}: {row_text[:80]}...")
                
                # 檢查是否包含契約相關關鍵字
                contract_keywords = ['TX', '臺指期', '台指期', '臺股期貨']
                has_contract = any(keyword in html_content for keyword in contract_keywords)
                print(f"   包含契約資料: {'是' if has_contract else '否'}")
                
            else:
                print(f"   請求失敗: {response.status_code}")
                
        except Exception as e:
            print(f"   發生錯誤: {str(e)}")

if __name__ == "__main__":
    test_taifex_response() 