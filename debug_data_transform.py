#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import sys
import os

# 添加當前目錄到路徑
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from taifex_crawler import TaifexCrawler, prepare_data_for_db

def test_data_transformation():
    """測試資料轉換過程"""
    print("🔍 測試資料轉換過程")
    
    # 1. 先爬取一筆資料
    crawler = TaifexCrawler(delay=1.0, max_retries=1)
    
    # 爬取今天的TX資料
    date_str = "2025/06/04"
    print(f"爬取 {date_str} TX 資料...")
    
    data = crawler.fetch_data(date_str, "TX", identity="自營商")
    if data:
        print(f"✅ 爬取成功: {data}")
        
        # 轉換為DataFrame
        df = pd.DataFrame([data])
        print(f"\n📊 原始DataFrame:")
        print(df)
        print(f"欄位: {list(df.columns)}")
        
        # 轉換為資料庫格式
        db_df = prepare_data_for_db(df)
        print(f"\n🗄️ 資料庫格式DataFrame:")
        print(db_df)
        print(f"欄位: {list(db_df.columns)}")
        
        # 檢查是否有空值
        print(f"\n🔍 檢查空值:")
        for col in db_df.columns:
            null_count = db_df[col].isnull().sum()
            empty_count = (db_df[col] == '').sum()
            print(f"  {col}: null={null_count}, empty={empty_count}")
            if null_count > 0 or empty_count > 0:
                print(f"    問題值: {db_df[col].tolist()}")
        
    else:
        print("❌ 爬取失敗")

if __name__ == "__main__":
    test_data_transformation() 