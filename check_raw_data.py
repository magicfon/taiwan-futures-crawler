#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import pandas as pd

conn = sqlite3.connect('data/taifex_data.db')

# 檢查資料表結構
print("📊 資料表結構:")
cursor = conn.cursor()
cursor.execute("PRAGMA table_info(futures_data)")
columns = cursor.fetchall()
for col in columns:
    print(f"   {col[1]}: {col[2]}")

# 檢查6/6的原始資料
print(f"\n📅 2025/06/06 原始資料:")
query = "SELECT * FROM futures_data WHERE date = '2025/06/06' ORDER BY id"
df = pd.read_sql_query(query, conn)
print(f"總共 {len(df)} 筆")
print(df[['date', 'contract_code', 'identity_type', 'position_type', 'long_position', 'short_position', 'net_position']])

conn.close() 