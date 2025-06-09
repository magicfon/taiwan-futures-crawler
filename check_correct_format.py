#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import pandas as pd

def check_database_format():
    print("🔍 檢查資料庫中的正確格式...")
    
    conn = sqlite3.connect('data/taifex_data.db')
    
    # 檢查6/6的資料
    query = """
    SELECT date, contract_code, identity_type, position_type, 
           long_position, short_position, net_position 
    FROM futures_data 
    WHERE date = '2025/06/06' 
    ORDER BY contract_code, identity_type
    """
    
    df = pd.read_sql_query(query, conn)
    
    print(f"📊 2025/06/06 資料庫內容 ({len(df)} 筆):")
    for _, row in df.iterrows():
        print(f"   {row['contract_code']} {row['identity_type']} {row['position_type']}: 多{row['long_position']}, 空{row['short_position']}, 淨{row['net_position']}")
    
    # 檢查是否有多方資料 (long_position > 0)
    long_data = df[df['long_position'] > 0]
    print(f"\n📈 有多方部位的記錄: {len(long_data)} 筆")
    
    # 檢查是否有空方資料 (short_position > 0)  
    short_data = df[df['short_position'] > 0]
    print(f"📉 有空方部位的記錄: {len(short_data)} 筆")
    
    # 檢查position_type的種類
    position_types = df['position_type'].unique()
    print(f"\n📋 position_type種類: {position_types}")
    
    conn.close()

if __name__ == "__main__":
    check_database_format() 