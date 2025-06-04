#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
修正後的台期所解析器
"""

import requests
from bs4 import BeautifulSoup
import re

def parse_number(text):
    """解析數字，處理千位分隔符"""
    if not text:
        return 0
    try:
        cleaned_str = re.sub(r'[^\d\-\.]', '', str(text))
        if '.' in cleaned_str:
            return float(cleaned_str) if cleaned_str else 0
        else:
            return int(cleaned_str) if cleaned_str else 0
    except:
        return 0

def fetch_and_parse_data(date_str, contract):
    """修正後的資料抓取和解析"""
    print(f"🔍 抓取 {date_str} 的 {contract} 資料...")
    
    BASE_URL = "https://www.taifex.com.tw/cht/3/futContractsDate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    params = {
        'queryType': '2',
        'marketCode': '0',
        'dateaddcnt': '',
        'commodity_id': contract,
        'queryDate': date_str
    }
    
    response = requests.get(BASE_URL, params=params, headers=headers, timeout=10)
    html_content = response.text
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 找到主要表格
    tables = soup.find_all('table')
    if not tables:
        print("❌ 沒有找到表格")
        return []
    
    main_table = max(tables, key=lambda t: len(str(t)))
    rows = main_table.find_all('tr')
    
    print(f"📊 表格有 {len(rows)} 行")
    
    results = []
    current_contract = None
    
    # 契約名稱對應
    contract_mapping = {
        'TX': '臺股期貨',
        'TE': '電子期貨', 
        'MTX': '小型臺指',
        'ZMX': '微型臺指',
        'NQF': '那斯達克'
    }
    
    target_contract_name = contract_mapping.get(contract, contract)
    
    for i, row in enumerate(rows):
        cells = row.find_all(['td', 'th'])
        if len(cells) < 3:
            continue
            
        # 取得所有單元格文字
        cell_texts = [cell.get_text(strip=True) for cell in cells]
        
        # 檢查是否是目標契約的起始行
        if len(cell_texts) >= 3 and target_contract_name in cell_texts[1]:
            current_contract = target_contract_name
            print(f"✅ 找到契約: {current_contract}")
            
            # 解析這行資料 (第一個身份別，通常是自營商)
            if len(cell_texts) >= 8:
                identity = cell_texts[2]  # 身份別
                data = {
                    '日期': date_str,
                    '契約名稱': contract,
                    '身份別': identity,
                    '多方交易口數': parse_number(cell_texts[3]),
                    '多方契約金額': parse_number(cell_texts[4]),
                    '空方交易口數': parse_number(cell_texts[5]),
                    '空方契約金額': parse_number(cell_texts[6]),
                    '多空淨額交易口數': parse_number(cell_texts[7]),
                    '多空淨額契約金額': parse_number(cell_texts[8]) if len(cell_texts) > 8 else 0,
                    '多方未平倉口數': parse_number(cell_texts[9]) if len(cell_texts) > 9 else 0,
                    '多方未平倉契約金額': parse_number(cell_texts[10]) if len(cell_texts) > 10 else 0,
                    '空方未平倉口數': parse_number(cell_texts[11]) if len(cell_texts) > 11 else 0,
                    '空方未平倉契約金額': parse_number(cell_texts[12]) if len(cell_texts) > 12 else 0,
                    '多空淨額未平倉口數': parse_number(cell_texts[13]) if len(cell_texts) > 13 else 0,
                    '多空淨額未平倉契約金額': parse_number(cell_texts[14]) if len(cell_texts) > 14 else 0
                }
                results.append(data)
                print(f"   {identity}: 多方 {data['多方交易口數']} 口, 空方 {data['空方交易口數']} 口")
                
        # 檢查是否是同契約的其他身份別行 (投信、外資)
        elif current_contract and len(cell_texts) >= 8:
            # 如果第一列不是數字（序號），可能是身份別資料
            if not cell_texts[0].isdigit() and cell_texts[0] in ['投信', '外資', '投顧']:
                identity = cell_texts[0]  # 身份別
                data = {
                    '日期': date_str,
                    '契約名稱': contract,
                    '身份別': identity,
                    '多方交易口數': parse_number(cell_texts[1]),
                    '多方契約金額': parse_number(cell_texts[2]),
                    '空方交易口數': parse_number(cell_texts[3]),
                    '空方契約金額': parse_number(cell_texts[4]),
                    '多空淨額交易口數': parse_number(cell_texts[5]),
                    '多空淨額契約金額': parse_number(cell_texts[6]),
                    '多方未平倉口數': parse_number(cell_texts[7]) if len(cell_texts) > 7 else 0,
                    '多方未平倉契約金額': parse_number(cell_texts[8]) if len(cell_texts) > 8 else 0,
                    '空方未平倉口數': parse_number(cell_texts[9]) if len(cell_texts) > 9 else 0,
                    '空方未平倉契約金額': parse_number(cell_texts[10]) if len(cell_texts) > 10 else 0,
                    '多空淨額未平倉口數': parse_number(cell_texts[11]) if len(cell_texts) > 11 else 0,
                    '多空淨額未平倉契約金額': parse_number(cell_texts[12]) if len(cell_texts) > 12 else 0
                }
                results.append(data)
                print(f"   {identity}: 多方 {data['多方交易口數']} 口, 空方 {data['空方交易口數']} 口")
            
            # 如果遇到下一個契約，重置current_contract
            elif cell_texts[0].isdigit() and len(cell_texts) > 1:
                for contract_name in contract_mapping.values():
                    if contract_name in cell_texts[1]:
                        current_contract = None  # 重置，表示已經到下一個契約
                        break
    
    print(f"✅ 共解析到 {len(results)} 筆資料")
    return results

# 測試
if __name__ == "__main__":
    # 測試TX契約
    data = fetch_and_parse_data("2024/12/31", "TX")
    
    print("\n📊 解析結果:")
    for item in data:
        print(f"  {item['身份別']}: 多方{item['多方交易口數']}口 空方{item['空方交易口數']}口 淨額{item['多空淨額交易口數']}口") 