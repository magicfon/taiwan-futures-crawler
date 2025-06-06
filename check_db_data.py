#!/usr/bin/env python3
"""
檢查資料庫資料格式
"""

from database_manager import TaifexDatabaseManager
import pandas as pd

def main():
    print("🔍 檢查資料庫資料格式...")
    
    db_manager = TaifexDatabaseManager()
    
    # 1. 檢查近期資料
    recent_data = db_manager.get_recent_data(30)
    print(f"📊 近30天資料: {len(recent_data)} 筆")
    
    if not recent_data.empty:
        print("\n📋 近期資料欄位:")
        for col in recent_data.columns:
            dtype = recent_data[col].dtype
            sample_value = recent_data[col].iloc[0] if len(recent_data) > 0 else None
            sample_type = type(sample_value)
            print(f"  - {col}: {dtype} (範例: {sample_value}, 類型: {sample_type})")
        
        print(f"\n📊 近期資料前3筆:")
        print(recent_data.head(3))
    
    # 2. 檢查摘要資料
    summary_data = db_manager.get_daily_summary(30)
    print(f"\n📈 摘要資料: {len(summary_data)} 筆")
    
    if not summary_data.empty:
        print("\n📋 摘要資料欄位:")
        for col in summary_data.columns:
            dtype = summary_data[col].dtype
            sample_value = summary_data[col].iloc[0] if len(summary_data) > 0 else None
            sample_type = type(sample_value)
            print(f"  - {col}: {dtype} (範例: {sample_value}, 類型: {sample_type})")
            
            # 特別檢查是否有bytes類型
            if any(isinstance(val, bytes) for val in summary_data[col].dropna()):
                print(f"    ⚠️ 發現bytes類型資料！")
                bytes_values = [val for val in summary_data[col].dropna() if isinstance(val, bytes)]
                print(f"    bytes範例: {bytes_values[:3]}")
        
        print(f"\n📊 摘要資料:")
        print(summary_data)
        
        # 檢查具體的問題欄位
        for col in summary_data.columns:
            if summary_data[col].dtype == 'object':
                print(f"\n🔍 檢查欄位 '{col}' 的資料類型:")
                for i, val in enumerate(summary_data[col]):
                    print(f"  [{i}] {type(val)}: {repr(val)}")

if __name__ == "__main__":
    main() 