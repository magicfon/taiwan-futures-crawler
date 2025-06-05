#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime
import sys
import os
import argparse

# 模擬完整的 parse_arguments 邏輯
def mock_parse_arguments():
    """模擬 parse_arguments 函數"""
    
    # 模擬命令行參數
    class MockArgs:
        def __init__(self):
            self.date_range = 'today'
            self.start_date = None
            self.end_date = None
            self.year = None
            self.month = None
    
    args = MockArgs()
    
    print("=" * 50)
    print("📋 模擬 parse_arguments 邏輯")
    print("=" * 50)
    print(f"輸入參數: date_range = '{args.date_range}'")
    
    # 處理新的日期範圍參數
    if args.date_range:
        print("\n🔍 進入 date_range 分支")
        
        if args.date_range.lower() == 'today':
            print("✅ 匹配到 'today' 分支")
            # 今天
            today = datetime.datetime.now()
            start_date = today
            end_date = today
            print(f"初始設定: start_date = {start_date}")
            print(f"初始設定: end_date = {end_date}")
            
    # 🔧 修復：統一時間部分 - 將所有日期設為當天的00:00:00，避免時間比較問題
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = end_date.replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"\n🔧 統一時間後:")
    print(f"start_date = {start_date}")
    print(f"end_date = {end_date}")
    
    # 確保結束日期不超過今天
    today_check = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"\n🕐 時間檢查:")
    print(f"end_date = {end_date}")
    print(f"today_check = {today_check}")
    print(f"end_date > today_check: {end_date > today_check}")
    
    if end_date > today_check:
        end_date = today_check
        print(f"⚠️ 調整後: end_date = {end_date}")
    else:
        print("✅ 日期檢查通過，無需調整")
    
    args.start_date = start_date
    args.end_date = end_date
    
    print(f"\n📋 最終結果:")
    print(f"start_date: {args.start_date}")
    print(f"end_date: {args.end_date}")
    
    return args

def mock_crawl_date_range(start_date, end_date):
    """模擬 crawl_date_range 中的日期判斷邏輯"""
    
    def _is_business_day(date):
        """檢查是否為交易日（非週末）"""
        day = date.weekday()
        return day < 5
    
    print("\n" + "=" * 50)
    print("📋 模擬 crawl_date_range 邏輯")
    print("=" * 50)
    
    # 調試：記錄系統時間信息
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"系統時間: {today.strftime('%Y/%m/%d %A')}")
    print(f"開始日期: {start_date.strftime('%Y/%m/%d %A')}")
    print(f"結束日期: {end_date.strftime('%Y/%m/%d %A')}")
    
    # 🔧 詳細檢查日期範圍計算
    days_diff = (end_date - start_date).days
    print(f"\n🔍 日期範圍計算:")
    print(f"(end_date - start_date).days = {days_diff}")
    print(f"range((end_date - start_date).days + 1) = range({days_diff + 1})")
    
    # 計算總任務數量
    date_range = [start_date + datetime.timedelta(days=x) for x in range((end_date - start_date).days + 1)]
    print(f"\n生成的日期範圍:")
    
    # 記錄所有日期和其交易日判斷
    for date in date_range:
        is_business = _is_business_day(date)
        print(f"  {date.strftime('%Y/%m/%d %A')} - 交易日: {is_business}")
    
    business_days = [d for d in date_range if _is_business_day(d)]
    
    print(f"\n📊 統計結果:")
    print(f"總日期數: {len(date_range)}")
    print(f"交易日數: {len(business_days)}")
    
    # 如果沒有交易日，直接返回空的DataFrame
    if len(business_days) == 0:
        print("❌ 判斷：在指定範圍內沒有找到交易日")
        return False
    else:
        print("✅ 判斷：找到交易日，開始爬取")
        return True

def main():
    print("🔍 完整流程調試工具 (修復版)")
    print("模擬：python taifex_crawler.py --date-range today")
    
    # 1. 模擬參數解析
    args = mock_parse_arguments()
    
    # 2. 模擬日期範圍處理
    success = mock_crawl_date_range(args.start_date, args.end_date)
    
    print("\n" + "=" * 50)
    print("🎯 問題分析")
    print("=" * 50)
    
    if success:
        print("✅ 邏輯正常，應該可以正常爬取資料")
    else:
        print("❌ 邏輯異常，這就是為什麼無法爬取的原因")
        
        # 進一步分析
        print("\n可能的問題:")
        print("1. 時間比較邏輯錯誤")
        print("2. 週末判斷邏輯錯誤") 
        print("3. 日期範圍計算錯誤")

if __name__ == "__main__":
    main() 