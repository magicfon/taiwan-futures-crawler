#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
檢查近30天交易日資料遺漏並自動補齊
在GitHub Actions每日爬取之前執行
"""

import sys
import os
import logging
from datetime import datetime, timedelta
import pandas as pd
import sqlite3
from database_manager import TaifexDatabaseManager
from google_sheets_manager import GoogleSheetsManager
import subprocess

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('missing_data_check.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class MissingDataChecker:
    """遺漏資料檢查和補齊工具"""
    
    def __init__(self):
        self.db = TaifexDatabaseManager()
        try:
            self.gm = GoogleSheetsManager()
            self.has_google_sheets = True
        except Exception as e:
            logger.warning(f"無法初始化Google Sheets: {e}")
            self.has_google_sheets = False
        
        # 預設契約代碼
        self.default_contracts = ['TX', 'TE', 'MTX']
        
    def is_trading_day(self, date):
        """檢查是否為交易日（非週末）"""
        return date.weekday() < 5  # 0-4 是週一到週五
    
    def get_trading_days_in_range(self, start_date, end_date):
        """取得指定範圍內的所有交易日"""
        trading_days = []
        current_date = start_date
        
        while current_date <= end_date:
            if self.is_trading_day(current_date):
                trading_days.append(current_date)
            current_date += timedelta(days=1)
        
        return trading_days
    
    def check_database_missing_dates(self, days=30):
        """檢查資料庫中遺漏的交易日"""
        logger.info(f"🔍 檢查近{days}天的交易日資料...")
        
        # 計算檢查範圍
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        # 取得所有應該有資料的交易日
        expected_trading_days = self.get_trading_days_in_range(start_date, end_date)
        logger.info(f"📅 期間內應有交易日: {len(expected_trading_days)}天")
        
        # 查詢資料庫中現有的日期
        conn = sqlite3.connect(self.db.db_path)
        query = """
        SELECT DISTINCT date 
        FROM futures_data 
        WHERE date >= ? AND date <= ?
        """
        
        existing_dates_df = pd.read_sql_query(
            query, 
            conn, 
            params=[
                start_date.strftime('%Y/%m/%d'),
                end_date.strftime('%Y/%m/%d')
            ]
        )
        conn.close()
        
        # 轉換現有日期為datetime物件
        existing_dates = set()
        for date_str in existing_dates_df['date']:
            try:
                # 處理不同的日期格式
                if '/' in date_str:
                    if len(date_str.split('/')[1]) == 1:  # 2025/6/5 格式
                        date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                    else:  # 2025/06/05 格式
                        date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                    existing_dates.add(date_obj.date())
            except ValueError as e:
                logger.warning(f"無法解析日期格式: {date_str} - {e}")
        
        # 找出遺漏的交易日
        missing_dates = []
        for trading_day in expected_trading_days:
            if trading_day.date() not in existing_dates:
                missing_dates.append(trading_day)
        
        logger.info(f"📊 資料庫現有交易日: {len(existing_dates)}天")
        logger.info(f"❌ 遺漏的交易日: {len(missing_dates)}天")
        
        if missing_dates:
            logger.info("遺漏的日期:")
            for missing_date in sorted(missing_dates):
                weekday_names = ['週一', '週二', '週三', '週四', '週五', '週六', '週日']
                weekday = weekday_names[missing_date.weekday()]
                logger.info(f"   📅 {missing_date.strftime('%Y/%m/%d')} ({weekday})")
        
        return missing_dates
    
    def crawl_missing_data(self, missing_dates):
        """爬取遺漏的資料"""
        if not missing_dates:
            logger.info("✅ 沒有遺漏的資料需要補齊")
            return True
        
        logger.info(f"🚀 開始補齊 {len(missing_dates)} 天的遺漏資料...")
        
        success_count = 0
        failed_dates = []
        
        for date in sorted(missing_dates):
            try:
                date_str = date.strftime('%Y-%m-%d')
                logger.info(f"📈 正在爬取 {date_str} 的資料...")
                
                # 執行爬蟲
                cmd = [
                    'python', 'taifex_crawler.py',
                    '--date-range', f'{date_str},{date_str}',
                    '--contracts', ','.join(self.default_contracts)
                ]
                
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    timeout=300  # 5分鐘超時
                )
                
                if result.returncode == 0:
                    logger.info(f"✅ {date_str} 資料爬取成功")
                    success_count += 1
                else:
                    logger.error(f"❌ {date_str} 資料爬取失敗: {result.stderr}")
                    failed_dates.append(date)
                
            except subprocess.TimeoutExpired:
                logger.error(f"❌ {date_str} 爬取超時")
                failed_dates.append(date)
            except Exception as e:
                logger.error(f"❌ {date_str} 爬取過程發生錯誤: {e}")
                failed_dates.append(date)
        
        logger.info(f"📊 補齊結果: 成功 {success_count} 天, 失敗 {len(failed_dates)} 天")
        
        if failed_dates:
            logger.warning("失敗的日期:")
            for failed_date in failed_dates:
                logger.warning(f"   ❌ {failed_date.strftime('%Y/%m/%d')}")
        
        return len(failed_dates) == 0
    
    def sync_to_google_sheets(self):
        """同步資料到Google Sheets"""
        if not self.has_google_sheets:
            logger.warning("⚠️ 跳過Google Sheets同步（認證未設定）")
            return True
        
        try:
            logger.info("📊 開始同步資料到Google Sheets...")
            
            # 取得最近30天的資料
            recent_data = self.db.get_recent_data(days=30)
            
            if recent_data.empty:
                logger.warning("⚠️ 沒有資料可以同步到Google Sheets")
                return False
            
            # 轉換為Google Sheets格式
            sheets_data = self.prepare_sheets_data(recent_data)
            
            # 上傳到Google Sheets
            success = self.gm.upload_data(sheets_data)
            
            if success:
                logger.info("✅ Google Sheets同步成功")
                return True
            else:
                logger.error("❌ Google Sheets同步失敗")
                return False
                
        except Exception as e:
            logger.error(f"❌ Google Sheets同步過程發生錯誤: {e}")
            return False
    
    def prepare_sheets_data(self, df):
        """準備Google Sheets格式的資料"""
        # 轉換為Google Sheets需要的格式
        sheets_data = []
        
        for _, row in df.iterrows():
            sheets_row = {
                '日期': row['date'],
                '契約名稱': row['contract_code'],
                '身份別': row['identity_type'],
                '多方交易口數': row['long_position'] if row['position_type'] == '多方' else 0,
                '空方交易口數': row['short_position'] if row['position_type'] == '空方' else 0,
                '多方未平倉口數': row['long_position'] if row['position_type'] == '多方' else 0,
                '空方未平倉口數': row['short_position'] if row['position_type'] == '空方' else 0,
                '多空淨額交易口數': row['net_position'],
                '多空淨額未平倉口數': row['net_position'],
                '更新時間': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            sheets_data.append(sheets_row)
        
        return pd.DataFrame(sheets_data)
    
    def run_complete_check(self, days=30):
        """執行完整的檢查和補齊流程"""
        logger.info("🚀 開始執行完整的資料檢查和補齊流程...")
        
        try:
            # 1. 檢查遺漏的交易日
            missing_dates = self.check_database_missing_dates(days)
            
            # 2. 補齊遺漏的資料
            if missing_dates:
                crawl_success = self.crawl_missing_data(missing_dates)
                if not crawl_success:
                    logger.warning("⚠️ 部分資料補齊失敗，但繼續執行")
            
            # 3. 同步到Google Sheets
            sync_success = self.sync_to_google_sheets()
            
            # 4. 生成報告
            self.generate_report(missing_dates, days)
            
            logger.info("✅ 完整檢查和補齊流程執行完成")
            return True
            
        except Exception as e:
            logger.error(f"❌ 檢查和補齊流程發生錯誤: {e}")
            return False
    
    def generate_report(self, missing_dates, days):
        """生成檢查報告"""
        logger.info("📝 生成檢查報告...")
        
        report = {
            'check_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'check_period_days': days,
            'missing_dates_count': len(missing_dates),
            'missing_dates': [d.strftime('%Y-%m-%d') for d in missing_dates],
            'database_status': 'healthy' if len(missing_dates) == 0 else 'has_gaps',
            'google_sheets_status': 'enabled' if self.has_google_sheets else 'disabled'
        }
        
        # 儲存報告
        import json
        os.makedirs('reports', exist_ok=True)
        
        with open('reports/missing_data_check_report.json', 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        
        logger.info(f"📋 報告已儲存到 reports/missing_data_check_report.json")
        
        # 在日誌中輸出摘要
        logger.info("=" * 50)
        logger.info("📊 檢查結果摘要:")
        logger.info(f"   檢查期間: 近{days}天")
        logger.info(f"   遺漏交易日: {len(missing_dates)}天")
        logger.info(f"   資料庫狀態: {report['database_status']}")
        logger.info(f"   Google Sheets: {report['google_sheets_status']}")
        logger.info("=" * 50)

def main():
    """主程式"""
    import argparse
    
    parser = argparse.ArgumentParser(description='檢查並補齊近期遺漏的交易日資料')
    parser.add_argument('--days', type=int, default=30, help='檢查天數 (預設: 30)')
    parser.add_argument('--skip-sync', action='store_true', help='跳過Google Sheets同步')
    
    args = parser.parse_args()
    
    logger.info("🔍 開始執行遺漏資料檢查和補齊...")
    
    try:
        checker = MissingDataChecker()
        
        if args.skip_sync:
            # 只檢查和補齊，不同步
            missing_dates = checker.check_database_missing_dates(args.days)
            if missing_dates:
                checker.crawl_missing_data(missing_dates)
            checker.generate_report(missing_dates, args.days)
        else:
            # 執行完整流程
            success = checker.run_complete_check(args.days)
            
            if success:
                logger.info("✅ 所有檢查和補齊任務完成")
                sys.exit(0)
            else:
                logger.error("❌ 檢查和補齊過程中發生錯誤")
                sys.exit(1)
                
    except Exception as e:
        logger.error(f"❌ 程式執行失敗: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 