#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
時間調試腳本 - 用於檢查GitHub Actions環境的時間設定
"""

import datetime
import os
import time

def main():
    print("=" * 50)
    print("時間調試資訊")
    print("=" * 50)
    
    # 1. Python datetime
    now = datetime.datetime.now()
    print(f"Python datetime.now(): {now}")
    print(f"格式化: {now.strftime('%Y/%m/%d %A %H:%M:%S')}")
    print(f"時區: {now.tzinfo}")
    
    # 2. UTC時間
    utc_now = datetime.datetime.utcnow()
    print(f"Python datetime.utcnow(): {utc_now}")
    print(f"UTC格式化: {utc_now.strftime('%Y/%m/%d %A %H:%M:%S')}")
    
    # 3. 系統環境變數
    print("\n環境變數:")
    time_vars = ['TZ', 'TIMEZONE', 'LC_TIME']
    for var in time_vars:
        value = os.environ.get(var, 'Not set')
        print(f"  {var}: {value}")
    
    # 4. 週幾判斷
    print(f"\n週幾判斷:")
    print(f"今天是: {now.strftime('%A')} (weekday={now.weekday()})")
    print(f"是否為週末: {now.weekday() >= 5}")
    print(f"是否為交易日: {now.weekday() < 5}")
    
    # 5. 測試特定日期
    test_dates = [
        "2024-01-15",  # 週一
        "2024-01-13",  # 週六  
        "2024-01-14",  # 週日
    ]
    
    print(f"\n測試日期:")
    for date_str in test_dates:
        test_date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        is_business = test_date.weekday() < 5
        print(f"  {date_str} ({test_date.strftime('%A')}): 交易日={is_business}")
    
    # 6. 今天範圍測試
    print(f"\n今天範圍測試:")
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"今天00:00: {today}")
    print(f"是否超過今天: {now > today}")
    
    print("=" * 50)

if __name__ == "__main__":
    main() 