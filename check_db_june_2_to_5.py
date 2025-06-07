#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
專門檢查資料庫中2025年6/2~6/5日期範圍的資料
"""

from database_manager import TaifexDatabaseManager
import pandas as pd
from datetime import datetime, timedelta
import sqlite3

def check_db_june_2_to_5():
    """檢查資料庫中2025年6/2~6/5的資料狀況"""
    try:
        print('📅 檢查資料庫中2025年6/2~6/5日期範圍的資料...')
        
        # 先檢查這些日期是星期幾
        target_dates = []
        for day in range(2, 6):  # 6/2 到 6/5
            date = datetime(2025, 6, day)
            weekday = date.weekday()  # 0=週一, 6=週日
            weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
            is_trading_day = weekday < 5  # 週一到週五是交易日
            
            target_dates.append({
                'date': date,
                'date_str': f"2025/06/{day:02d}",
                'date_str_short': f"2025/6/{day}",
                'weekday': weekday_names[weekday],
                'is_trading_day': is_trading_day
            })
            
            print(f"📆 2025年6月{day}日是 {weekday_names[weekday]} {'✅交易日' if is_trading_day else '❌非交易日'}")
        
        # 檢查資料庫
        print(f"\n🔍 檢查資料庫資料...")
        db = TaifexDatabaseManager()
        
        # 直接使用SQLite連接來執行查詢
        conn = sqlite3.connect(db.db_path)
        
        # 檢查詳細資料表
        print(f"\n📊 檢查詳細資料表...")
        query = """
        SELECT DISTINCT date, COUNT(*) as count
        FROM futures_data 
        WHERE date LIKE '2025/06/%' OR date LIKE '2025/6/%'
        GROUP BY date
        ORDER BY date
        """
        
        df_details = pd.read_sql_query(query, conn)
        print(f"📋 6月份詳細資料:")
        if not df_details.empty:
            for _, row in df_details.iterrows():
                date_str = row['date']
                count = row['count']
                
                # 檢查是否為目標日期
                is_target = any(
                    target['date_str'] == date_str or target['date_str_short'] == date_str 
                    for target in target_dates
                )
                
                status = "🎯" if is_target else "📅"
                print(f"   {status} {date_str}: {count}筆")
        else:
            print("   ❌ 沒有6月份的詳細資料")
        
        # 檢查摘要資料表
        print(f"\n📊 檢查摘要資料表...")
        query_summary = """
        SELECT date, total_contracts, total_volume, foreign_net, dealer_net, trust_net, data_status, created_at
        FROM daily_summary 
        WHERE date LIKE '2025/06/%' OR date LIKE '2025/6/%'
        ORDER BY date
        """
        
        df_summary = pd.read_sql_query(query_summary, conn)
        print(f"📋 6月份摘要資料:")
        if not df_summary.empty:
            for _, row in df_summary.iterrows():
                date_str = row['date']
                status = row['data_status']
                volume = row['total_volume']
                created = row['created_at']
                
                # 檢查是否為目標日期
                is_target = any(
                    target['date_str'] == date_str or target['date_str_short'] == date_str 
                    for target in target_dates
                )
                
                status_icon = "🎯" if is_target else "📅"
                print(f"   {status_icon} {date_str}: 狀態={status}, 成交量={volume}, 建立時間={created}")
        else:
            print("   ❌ 沒有6月份的摘要資料")
        
        # 針對每個目標日期進行詳細檢查
        print(f"\n🔍 針對目標日期進行詳細檢查...")
        for target in target_dates:
            date_str = target['date_str']
            date_str_short = target['date_str_short']
            weekday = target['weekday']
            
            print(f"\n📅 檢查 {date_str_short} ({weekday}):")
            
            # 檢查詳細資料
            query_detail = """
            SELECT COUNT(*) as count, 
                   GROUP_CONCAT(DISTINCT contract_code) as contracts,
                   GROUP_CONCAT(DISTINCT identity_type) as identities
            FROM futures_data 
            WHERE date = ? OR date = ?
            """
            
            df_detail = pd.read_sql_query(query_detail, conn, params=[date_str, date_str_short])
            
            if not df_detail.empty and df_detail.iloc[0]['count'] > 0:
                count = df_detail.iloc[0]['count']
                contracts = df_detail.iloc[0]['contracts']
                identities = df_detail.iloc[0]['identities']
                print(f"   ✅ 詳細資料: {count}筆")
                print(f"   📋 契約: {contracts}")
                print(f"   👥 身份: {identities}")
            else:
                print(f"   ❌ 沒有詳細資料")
            
            # 檢查摘要資料
            query_summary_detail = """
            SELECT * FROM daily_summary 
            WHERE date = ? OR date = ?
            """
            
            df_summary_detail = pd.read_sql_query(query_summary_detail, conn, params=[date_str, date_str_short])
            
            if not df_summary_detail.empty:
                row = df_summary_detail.iloc[0]
                print(f"   ✅ 摘要資料: 狀態={row['data_status']}, 成交量={row['total_volume']}")
                print(f"   📊 三大法人淨額: 外資={row['foreign_net']}, 投信={row['trust_net']}, 自營={row['dealer_net']}")
            else:
                print(f"   ❌ 沒有摘要資料")
        
        # 檢查最近的資料更新時間
        print(f"\n🕐 檢查最近的資料更新時間...")
        query_latest = """
        SELECT MAX(created_at) as latest_update, COUNT(*) as total_records
        FROM futures_data
        """
        
        df_latest = pd.read_sql_query(query_latest, conn)
        if not df_latest.empty:
            latest_update = df_latest.iloc[0]['latest_update']
            total_records = df_latest.iloc[0]['total_records']
            print(f"📊 資料庫總記錄數: {total_records}")
            print(f"🕐 最後更新時間: {latest_update}")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f'❌ 檢查失敗: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    check_db_june_2_to_5() 