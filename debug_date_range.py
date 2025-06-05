#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime
import sys
import os

# 添加當前目錄到路徑
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_business_day_logic():
    """測試交易日判斷邏輯"""
    def _is_business_day(date):
        """檢查是否為交易日（非週末）"""
        day = date.weekday()
        # 0-4 是週一到週五，5-6 是週末
        return day < 5
    
    # 測試今天和本週的日期
    today = datetime.datetime(2025, 6, 4)
    print(f"測試日期: {today.strftime('%Y/%m/%d %A')}")
    print(f"weekday(): {today.weekday()}")
    print(f"是交易日: {_is_business_day(today)}")
    
    print("\n本週所有日期測試:")
    for i in range(-3, 4):  # 本週前後幾天
        test_date = today + datetime.timedelta(days=i)
        is_business = _is_business_day(test_date)
        print(f"{test_date.strftime('%Y/%m/%d %A')} (weekday={test_date.weekday()}) - 交易日: {is_business}")

def test_date_range_logic():
    """測試日期範圍邏輯"""
    def _is_business_day(date):
        """檢查是否為交易日（非週末）"""
        day = date.weekday()
        return day < 5
    
    start_date = datetime.datetime(2025, 6, 4)
    end_date = datetime.datetime(2025, 6, 4)
    
    print(f"\n測試日期範圍: {start_date.strftime('%Y/%m/%d')} - {end_date.strftime('%Y/%m/%d')}")
    
    # 模擬原始邏輯
    date_range = [start_date + datetime.timedelta(days=x) for x in range((end_date - start_date).days + 1)]
    print(f"生成的日期範圍: {[d.strftime('%Y/%m/%d %A') for d in date_range]}")
    
    business_days = [d for d in date_range if _is_business_day(d)]
    print(f"篩選出的交易日: {[d.strftime('%Y/%m/%d %A') for d in business_days]}")
    
    print(f"交易日數量: {len(business_days)}")
    
    if len(business_days) == 0:
        print("❌ 判斷：沒有交易日")
    else:
        print("✅ 判斷：有交易日")

def test_today_parsing():
    """測試 'today' 參數解析邏輯"""
    # 模擬 parse_arguments 中的邏輯
    print("\n測試 'today' 參數解析:")
    
    # 模擬 date_range = 'today'
    date_range = 'today'
    
    if date_range.lower() == 'today':
        # 今天
        today = datetime.datetime.now()
        start_date = today
        end_date = today
        print(f"解析結果: start_date={start_date.strftime('%Y/%m/%d %A')}, end_date={end_date.strftime('%Y/%m/%d %A')}")
    
    # 確保結束日期不超過今天
    today_check = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if end_date > today_check:
        end_date = today_check
        print(f"調整後的 end_date: {end_date.strftime('%Y/%m/%d %A')}")
    
    print(f"最終日期範圍: {start_date.strftime('%Y/%m/%d')} - {end_date.strftime('%Y/%m/%d')}")

def main():
    print("=" * 60)
    print("🐛 日期範圍調試工具")
    print("=" * 60)
    
    # 1. 測試交易日判斷邏輯
    test_business_day_logic()
    
    # 2. 測試日期範圍邏輯  
    test_date_range_logic()
    
    # 3. 測試 'today' 解析
    test_today_parsing()

if __name__ == "__main__":
    main() 