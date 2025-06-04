#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試Telegram圖表發送功能
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging
import os

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("測試Telegram圖表")

def create_test_data():
    """創建測試資料"""
    logger.info("🔧 創建測試資料...")
    
    # 生成30天的日期範圍
    end_date = datetime.now()
    start_date = end_date - timedelta(days=30)
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # 過濾工作日
    business_dates = [d for d in dates if d.weekday() < 5]
    
    test_data = []
    contracts = ['TX', 'TE', 'MTX']
    identities = ['自營商', '投信', '外資']
    
    for date in business_dates:
        for contract in contracts:
            for identity in identities:
                # 生成模擬資料
                test_data.append({
                    '日期': date,
                    '契約名稱': contract,
                    '身份別': identity,
                    '多空淨額交易口數': np.random.randint(-5000, 5000),
                    '多空淨額未平倉口數': np.random.randint(-15000, 15000)
                })
    
    df = pd.DataFrame(test_data)
    logger.info(f"✅ 已創建 {len(df)} 筆測試資料")
    return df

def test_chart_generation():
    """測試圖表生成"""
    try:
        from chart_generator import ChartGenerator
        
        logger.info("📊 測試圖表生成功能...")
        
        # 創建測試資料
        df = create_test_data()
        
        # 初始化圖表生成器
        chart_generator = ChartGenerator(output_dir="test_charts")
        
        # 生成圖表
        chart_paths = chart_generator.generate_all_charts(df)
        summary_text = chart_generator.generate_summary_text(df)
        
        logger.info(f"✅ 已生成 {len(chart_paths)} 個圖表:")
        for path in chart_paths:
            logger.info(f"  📈 {os.path.basename(path)}")
        
        logger.info("📝 摘要文字:")
        print(summary_text)
        
        return chart_paths, summary_text
        
    except ImportError:
        logger.error("❌ 圖表生成模組未找到，請安裝 matplotlib")
        return [], ""
    except Exception as e:
        logger.error(f"❌ 圖表生成失敗: {e}")
        return [], ""

def test_telegram_notification(chart_paths, summary_text):
    """測試Telegram通知"""
    try:
        from telegram_notifier import TelegramNotifier
        
        logger.info("📱 測試Telegram通知功能...")
        
        # 初始化Telegram通知器
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        # 測試連線
        if not notifier.test_connection():
            logger.error("❌ Telegram連線失敗")
            return False
        
        # 發送測試訊息
        test_message = "🧪 台期所爬蟲系統測試\n📊 圖表生成和發送功能測試"
        if notifier.send_message(test_message):
            logger.info("✅ 測試訊息發送成功")
        
        # 如果有圖表，發送圖表
        if chart_paths:
            success = notifier.send_chart_report(chart_paths, summary_text)
            if success:
                logger.info("✅ 圖表報告發送成功")
                return True
            else:
                logger.warning("⚠️ 圖表報告發送部分失敗")
                return False
        else:
            logger.warning("⚠️ 沒有圖表可發送")
            return True
            
    except ImportError:
        logger.error("❌ Telegram通知模組未找到")
        return False
    except Exception as e:
        logger.error(f"❌ Telegram通知失敗: {e}")
        return False

def main():
    """主測試程式"""
    logger.info("=== 台期所Telegram圖表發送功能測試 ===\n")
    
    # 1. 測試圖表生成
    chart_paths, summary_text = test_chart_generation()
    
    if not chart_paths:
        logger.error("❌ 圖表生成失敗，無法繼續測試")
        return 1
    
    # 2. 測試Telegram發送
    telegram_success = test_telegram_notification(chart_paths, summary_text)
    
    # 3. 清理測試檔案（可選）
    cleanup = input("\n🗑️ 是否刪除測試圖表檔案？ (y/N): ").lower().strip()
    if cleanup == 'y':
        import shutil
        if os.path.exists("test_charts"):
            shutil.rmtree("test_charts")
            logger.info("✅ 測試檔案已清理")
    else:
        logger.info("📁 測試圖表保留在 test_charts/ 目錄中")
    
    # 4. 總結
    logger.info("\n=== 測試結果摘要 ===")
    logger.info(f"📊 圖表生成: {'✅ 成功' if chart_paths else '❌ 失敗'}")
    logger.info(f"📱 Telegram發送: {'✅ 成功' if telegram_success else '❌ 失敗'}")
    
    if chart_paths and telegram_success:
        logger.info("🎉 所有功能測試通過！")
        logger.info("💡 現在您可以執行主程式來爬取真實資料並自動發送圖表了")
        return 0
    else:
        logger.error("❌ 部分功能測試失敗，請檢查配置")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 