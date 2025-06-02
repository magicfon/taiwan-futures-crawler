#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
台期所資料庫管理系統
支援SQLite本地資料庫和雲端資料庫的資料管理
"""

import sqlite3
import pandas as pd
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import logging

class TaifexDatabaseManager:
    """台期所資料庫管理器"""
    
    def __init__(self, db_path="data/taifex_data.db"):
        """初始化資料庫管理器"""
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(exist_ok=True)
        self.logger = logging.getLogger(__name__)
        self.init_database()
    
    def init_database(self):
        """初始化資料庫結構"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 建立主要資料表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS futures_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                contract_code TEXT NOT NULL,
                identity_type TEXT NOT NULL,
                position_type TEXT NOT NULL,
                long_position INTEGER,
                short_position INTEGER,
                net_position INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(date, contract_code, identity_type, position_type)
            )
        ''')
        
        # 建立日報摘要表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS daily_summary (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL UNIQUE,
                total_contracts INTEGER,
                total_volume INTEGER,
                foreign_net INTEGER,
                dealer_net INTEGER,
                trust_net INTEGER,
                data_status TEXT DEFAULT 'complete',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 建立索引提升查詢效能
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_date ON futures_data(date)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_contract ON futures_data(contract_code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_identity ON futures_data(identity_type)')
        
        conn.commit()
        conn.close()
        self.logger.info(f"資料庫初始化完成：{self.db_path}")
    
    def insert_data(self, df):
        """插入或更新資料"""
        conn = sqlite3.connect(self.db_path)
        
        try:
            # 使用replace into 來處理重複資料
            df.to_sql('futures_data', conn, if_exists='append', index=False, method='replace')
            
            # 更新日報摘要
            self.update_daily_summary(df)
            
            conn.commit()
            self.logger.info(f"成功插入 {len(df)} 筆資料")
            
        except Exception as e:
            conn.rollback()
            self.logger.error(f"資料插入失敗：{e}")
            raise
        finally:
            conn.close()
    
    def update_daily_summary(self, df):
        """更新每日摘要"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 按日期分組計算摘要
        for date in df['date'].unique():
            day_data = df[df['date'] == date]
            
            # 計算三大法人淨部位
            foreign_net = day_data[day_data['identity_type'] == '外資']['net_position'].sum()
            dealer_net = day_data[day_data['identity_type'] == '自營商']['net_position'].sum()
            trust_net = day_data[day_data['identity_type'] == '投信']['net_position'].sum()
            
            cursor.execute('''
                INSERT OR REPLACE INTO daily_summary 
                (date, total_contracts, total_volume, foreign_net, dealer_net, trust_net, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ''', (
                date,
                len(day_data['contract_code'].unique()),
                day_data['long_position'].sum() + day_data['short_position'].sum(),
                foreign_net,
                dealer_net,
                trust_net
            ))
        
        conn.commit()
        conn.close()
    
    def get_recent_data(self, days=30):
        """取得最近N天的資料"""
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT * FROM futures_data 
            WHERE date >= ? AND date <= ?
            ORDER BY date DESC, contract_code, identity_type
        '''
        
        df = pd.read_sql_query(query, conn, params=[
            start_date.strftime('%Y/%m/%d'),
            end_date.strftime('%Y/%m/%d')
        ])
        
        conn.close()
        return df
    
    def get_daily_summary(self, days=30):
        """取得每日摘要"""
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        conn = sqlite3.connect(self.db_path)
        
        query = '''
            SELECT * FROM daily_summary 
            WHERE date >= ? AND date <= ?
            ORDER BY date DESC
        '''
        
        df = pd.read_sql_query(query, conn, params=[
            start_date.strftime('%Y/%m/%d'),
            end_date.strftime('%Y/%m/%d')
        ])
        
        conn.close()
        return df
    
    def export_to_excel(self, output_path, days=30):
        """匯出資料到Excel"""
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 匯出原始資料
            recent_data = self.get_recent_data(days)
            recent_data.to_excel(writer, sheet_name='原始資料', index=False)
            
            # 匯出每日摘要
            daily_summary = self.get_daily_summary(days)
            daily_summary.to_excel(writer, sheet_name='每日摘要', index=False)
            
            # 匯出三大法人趨勢
            self.export_institutional_trends(writer, days)
        
        self.logger.info(f"資料匯出完成：{output_path}")
    
    def export_institutional_trends(self, writer, days=30):
        """匯出三大法人趨勢分析"""
        conn = sqlite3.connect(self.db_path)
        
        # 三大法人每日淨部位趨勢
        query = '''
            SELECT 
                date,
                SUM(CASE WHEN identity_type = '外資' THEN net_position ELSE 0 END) as 外資淨部位,
                SUM(CASE WHEN identity_type = '自營商' THEN net_position ELSE 0 END) as 自營商淨部位,
                SUM(CASE WHEN identity_type = '投信' THEN net_position ELSE 0 END) as 投信淨部位
            FROM futures_data
            WHERE date >= date('now', '-{} days')
            GROUP BY date
            ORDER BY date DESC
        '''.format(days)
        
        trends_df = pd.read_sql_query(query, conn)
        trends_df.to_excel(writer, sheet_name='三大法人趨勢', index=False)
        
        conn.close()
    
    def backup_to_csv(self, backup_dir="backup"):
        """備份資料庫到CSV"""
        backup_path = Path(backup_dir)
        backup_path.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        conn = sqlite3.connect(self.db_path)
        
        # 備份主要資料
        main_data = pd.read_sql_query('SELECT * FROM futures_data', conn)
        main_data.to_csv(backup_path / f'futures_data_{timestamp}.csv', index=False)
        
        # 備份摘要資料
        summary_data = pd.read_sql_query('SELECT * FROM daily_summary', conn)
        summary_data.to_csv(backup_path / f'daily_summary_{timestamp}.csv', index=False)
        
        conn.close()
        self.logger.info(f"資料備份完成：{backup_path}")


class CloudDatabaseManager:
    """雲端資料庫管理器（支援Supabase）"""
    
    def __init__(self, config_file="config/cloud_db.json"):
        """初始化雲端資料庫連接"""
        self.config_file = Path(config_file)
        self.config = self.load_config()
        self.logger = logging.getLogger(__name__)
    
    def load_config(self):
        """載入雲端資料庫設定"""
        if self.config_file.exists():
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # 建立預設設定檔
            default_config = {
                "supabase": {
                    "url": "your-supabase-url",
                    "key": "your-supabase-anon-key",
                    "enabled": False
                },
                "backup_enabled": True,
                "sync_interval_hours": 24
            }
            
            self.config_file.parent.mkdir(exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            
            return default_config
    
    def sync_to_cloud(self, local_db_manager):
        """同步本地資料到雲端"""
        if not self.config.get('supabase', {}).get('enabled', False):
            self.logger.info("雲端同步已停用")
            return
        
        try:
            # 這裡可以實作Supabase同步邏輯
            self.logger.info("雲端同步功能待實作")
        except Exception as e:
            self.logger.error(f"雲端同步失敗：{e}")


if __name__ == "__main__":
    # 測試資料庫管理器
    db_manager = TaifexDatabaseManager()
    
    # 範例資料
    sample_data = pd.DataFrame({
        'date': ['2024/01/01', '2024/01/01'],
        'contract_code': ['TX', 'TX'],
        'identity_type': ['外資', '自營商'],
        'position_type': ['多方', '多方'],
        'long_position': [1000, 500],
        'short_position': [800, 600],
        'net_position': [200, -100]
    })
    
    db_manager.insert_data(sample_data)
    print("資料庫測試完成！") 