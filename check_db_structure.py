#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import os

def check_database():
    db_path = "data/taifex_data.db"
    
    if not os.path.exists(db_path):
        print("❌ 資料庫檔案不存在")
        return
    
    print(f"✅ 資料庫檔案存在: {db_path}")
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # 檢查表結構
    print("\n📋 資料庫表結構:")
    cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='futures_data'")
    result = cursor.fetchone()
    
    if result:
        print(result[0])
    else:
        print("❌ futures_data 表不存在")
    
    # 檢查資料筆數
    print("\n📊 資料統計:")
    try:
        cursor.execute("SELECT COUNT(*) FROM futures_data")
        total_count = cursor.fetchone()[0]
        print(f"總筆數: {total_count}")
        
        if total_count > 0:
            # 檢查日期範圍
            cursor.execute("SELECT MIN(date), MAX(date) FROM futures_data")
            min_date, max_date = cursor.fetchone()
            print(f"日期範圍: {min_date} ~ {max_date}")
            
            # 檢查各契約數量
            cursor.execute("SELECT contract_code, COUNT(*) FROM futures_data GROUP BY contract_code")
            contracts = cursor.fetchall()
            print("\n各契約資料筆數:")
            for contract, count in contracts:
                print(f"  {contract}: {count}")
                
            # 檢查最近幾筆資料
            cursor.execute("SELECT * FROM futures_data ORDER BY date DESC LIMIT 5")
            recent_data = cursor.fetchall()
            print("\n最近5筆資料:")
            for row in recent_data:
                print(f"  {row}")
        
    except Exception as e:
        print(f"❌ 查詢錯誤: {e}")
    
    conn.close()

if __name__ == "__main__":
    check_database() 