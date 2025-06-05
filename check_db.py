#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from database_manager import TaifexDatabaseManager
import pandas as pd

def main():
    db = TaifexDatabaseManager()
    
    # 檢查全部資料
    all_data = db.get_recent_data(3650)  # 10年
    print(f"資料庫總筆數: {len(all_data)}")
    
    if not all_data.empty:
        print(f"最新日期: {all_data['date'].max()}")
        print(f"最舊日期: {all_data['date'].min()}")
        print(f"日期數量: {all_data['date'].nunique()}")
        
        # 檢查最近30天
        recent_30 = db.get_recent_data(30)
        print(f"最近30天資料筆數: {len(recent_30)}")
        
        # 檢查最近60天
        recent_60 = db.get_recent_data(60)
        print(f"最近60天資料筆數: {len(recent_60)}")
        
        # 檢查各契約資料
        if 'contract_code' in all_data.columns:
            contracts = all_data['contract_code'].value_counts()
            print("\n各契約資料筆數:")
            for contract, count in contracts.items():
                print(f"  {contract}: {count}")
    else:
        print("資料庫是空的")

if __name__ == "__main__":
    main() 