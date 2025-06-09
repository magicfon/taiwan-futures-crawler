#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
台期所每日自動爬取排程
支援兩階段資料爬取：
1. 下午2點：交易量資料
2. 下午3點半：完整資料（包含未平倉）
"""

import subprocess
import sys
import logging
import datetime
import time
import argparse
from pathlib import Path
import schedule

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("daily_schedule.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("每日排程")

class TaifexScheduler:
    """台期所資料爬取排程器"""
    
    def __init__(self, contracts=['TX', 'TE', 'MTX'], identities=['ALL']):
        """
        初始化排程器
        
        Args:
            contracts: 要爬取的契約列表
            identities: 要爬取的身份別列表
        """
        self.contracts = contracts
        self.identities = identities
        self.script_path = Path(__file__).parent / "taifex_crawler.py"
        
    def run_crawler(self, data_type='COMPLETE', test_mode=False):
        """
        執行爬蟲
        
        Args:
            data_type: 資料類型 ('TRADING' 或 'COMPLETE')
            test_mode: 測試模式，不實際執行只顯示命令
        """
        # 構建命令
        cmd = [
            sys.executable,
            str(self.script_path),
            '--date-range', 'today',
            '--contracts', ','.join(self.contracts),
            '--identities'] + self.identities + [
            '--data_type', data_type,
            '--max_workers', '5',  # 降低並發數避免被封IP
            '--delay', '1.0'       # 增加延遲
        ]
        
        if test_mode:
            logger.info(f"測試模式 - 將執行命令: {' '.join(cmd)}")
            return True
        
        try:
            logger.info(f"🚀 開始爬取{data_type}資料...")
            
            # 執行爬蟲
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,  # 5分鐘超時
                encoding='utf-8'
            )
            
            if result.returncode == 0:
                logger.info(f"✅ {data_type}資料爬取成功")
                if result.stdout:
                    logger.info("程式輸出:")
                    for line in result.stdout.split('\n'):
                        if line.strip():
                            logger.info(f"  {line}")
                return True
            else:
                logger.error(f"❌ {data_type}資料爬取失敗 (退出碼: {result.returncode})")
                if result.stderr:
                    logger.error("錯誤訊息:")
                    for line in result.stderr.split('\n'):
                        if line.strip():
                            logger.error(f"  {line}")
                return False
                
        except subprocess.TimeoutExpired:
            logger.error(f"❌ {data_type}資料爬取超時")
            return False
        except Exception as e:
            logger.error(f"❌ 執行{data_type}爬蟲時發生錯誤: {e}")
            return False
    
    def run_trading_data(self):
        """執行交易量資料爬取（下午2點）"""
        logger.info("📊 執行下午2點交易量資料爬取")
        return self.run_crawler('TRADING')
    
    def run_complete_data(self):
        """執行完整資料爬取（下午3點半）"""
        logger.info("📈 執行下午3點半完整資料爬取")
        return self.run_crawler('COMPLETE')
    
    def setup_schedule(self):
        """設定每日排程"""
        # 下午2點爬取交易量資料
        schedule.every().monday.at("14:00").do(self.run_trading_data)
        schedule.every().tuesday.at("14:00").do(self.run_trading_data)
        schedule.every().wednesday.at("14:00").do(self.run_trading_data)
        schedule.every().thursday.at("14:00").do(self.run_trading_data)
        schedule.every().friday.at("14:00").do(self.run_trading_data)
        
        # 下午3點半爬取完整資料
        schedule.every().monday.at("15:30").do(self.run_complete_data)
        schedule.every().tuesday.at("15:30").do(self.run_complete_data)
        schedule.every().wednesday.at("15:30").do(self.run_complete_data)
        schedule.every().thursday.at("15:30").do(self.run_complete_data)
        schedule.every().friday.at("15:30").do(self.run_complete_data)
        
        logger.info("⏰ 已設定每日排程:")
        logger.info("  - 週一到週五 14:00: 爬取交易量資料")
        logger.info("  - 週一到週五 15:30: 爬取完整資料")
    
    def run_now(self, data_type):
        """立即執行指定類型的爬取"""
        current_time = datetime.datetime.now()
        logger.info(f"🕐 立即執行模式 ({current_time.strftime('%H:%M')})")
        
        if data_type.upper() == 'TRADING':
            return self.run_trading_data()
        elif data_type.upper() == 'COMPLETE':
            return self.run_complete_data()
        else:
            logger.error(f"❌ 不支援的資料類型: {data_type}")
            return False
    
    def start_scheduler(self):
        """啟動排程器"""
        self.setup_schedule()
        logger.info("🤖 台期所資料爬取排程器已啟動")
        logger.info("按 Ctrl+C 停止排程器")
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(60)  # 每分鐘檢查一次
        except KeyboardInterrupt:
            logger.info("⏹️ 使用者中斷，排程器已停止")

def main():
    """主程序"""
    parser = argparse.ArgumentParser(description='台期所每日自動爬取排程')
    
    parser.add_argument('--mode', type=str, choices=['schedule', 'now', 'test'],
                        default='schedule', help='執行模式')
    parser.add_argument('--data_type', type=str, choices=['TRADING', 'COMPLETE'],
                        help='立即執行時的資料類型')
    parser.add_argument('--contracts', type=str, default='TX,TE,MTX',
                        help='要爬取的契約，用逗號分隔')
    parser.add_argument('--identities', type=str, default='ALL',
                        help='要爬取的身份別')
    
    args = parser.parse_args()
    
    # 解析契約
    contracts = [c.strip() for c in args.contracts.split(',')]
    
    # 解析身份別
    if args.identities.upper() == 'ALL':
        identities = ['ALL']
    else:
        identities = [i.strip() for i in args.identities.split(',')]
    
    # 創建排程器
    scheduler = TaifexScheduler(contracts=contracts, identities=identities)
    
    if args.mode == 'schedule':
        # 排程模式
        scheduler.start_scheduler()
    elif args.mode == 'now':
        # 立即執行模式
        if not args.data_type:
            current_hour = datetime.datetime.now().hour
            if 14 <= current_hour < 15:
                data_type = 'TRADING'
                logger.info("🕐 當前時間適合爬取交易量資料")
            else:
                data_type = 'COMPLETE'
                logger.info("🕐 當前時間適合爬取完整資料")
        else:
            data_type = args.data_type
        
        success = scheduler.run_now(data_type)
        if success:
            logger.info("✅ 爬取任務完成")
            sys.exit(0)
        else:
            logger.error("❌ 爬取任務失敗")
            sys.exit(1)
    elif args.mode == 'test':
        # 測試模式
        logger.info("🧪 測試模式 - 檢查命令構建")
        scheduler.run_crawler('TRADING', test_mode=True)
        scheduler.run_crawler('COMPLETE', test_mode=True)

if __name__ == "__main__":
    main() 