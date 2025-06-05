#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime
import sys
import os

# 添加當前目錄到路徑
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from taifex_crawler import TaifexCrawler

def main():
    # 檢查今天的日期信息
    today = datetime.datetime(2025, 6, 4)
    print(f"檢查日期: {today.strftime('%Y/%m/%d %A')}")
    print(f"星期幾: {today.weekday()} (0=週一, 6=週日)")
    print(f"是交易日: {today.weekday() < 5}")
    
    # 檢查系統當前時間
    now = datetime.datetime.now()
    print(f"系統當前時間: {now.strftime('%Y/%m/%d %H:%M:%S %A')}")
    
    # 嘗試爬取今天的資料
    print("\n開始測試爬取今天的資料...")
    
    crawler = TaifexCrawler(output_dir="test_output", delay=1.0, max_retries=2)
    
    # 測試爬取今天的TX資料
    date_str = "2025/06/04"
    print(f"測試爬取 {date_str} TX 資料...")
    
    try:
        data = crawler.fetch_data(date_str, "TX")
        if data:
            print("✅ 成功爬取到資料:")
            for key, value in data.items():
                print(f"  {key}: {value}")
        else:
            print("❌ 沒有爬取到資料")
    except Exception as e:
        print(f"❌ 爬取失敗: {e}")

if __name__ == "__main__":
    main() 