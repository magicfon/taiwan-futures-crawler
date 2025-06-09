#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
使用真實台期所資料測試Telegram圖表發送功能
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging
import os
from chart_generator import ChartGenerator
from telegram_notifier import TelegramNotifier

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("真實資料Telegram測試")

def load_latest_data():
    """載入最新的台期所資料"""
    logger.info("📊 載入最新台期所資料...")
    
    # 尋找最新的資料檔案
    data_file = 'output/taifex_20250101-20250608_TX_TE_MTX_ZMX_NQF_自營商_投信_外資.csv'
    
    if not os.path.exists(data_file):
        logger.error(f"❌ 資料檔案不存在: {data_file}")
        return None
    
    try:
        df = pd.read_csv(data_file)
        logger.info(f"✅ 成功載入 {len(df)} 筆資料")
        logger.info(f"📅 資料期間: {df['日期'].min()} ~ {df['日期'].max()}")
        
        # 轉換日期格式並排序
        df['日期'] = pd.to_datetime(df['日期'])
        df = df.sort_values('日期')
        
        # 只取最近30天的資料
        end_date = df['日期'].max()
        start_date = end_date - timedelta(days=30)
        df_recent = df[df['日期'] >= start_date].copy()
        
        logger.info(f"📊 最近30天資料: {len(df_recent)} 筆")
        logger.info(f"🎯 涵蓋契約: {', '.join(df_recent['契約名稱'].unique())}")
        logger.info(f"👥 涵蓋身份: {', '.join(df_recent['身份別'].unique())}")
        
        return df_recent
        
    except Exception as e:
        logger.error(f"❌ 載入資料失敗: {e}")
        return None

def generate_charts_with_real_data():
    """使用真實資料生成圖表"""
    logger.info("📈 使用真實資料生成圖表...")
    
    # 載入資料
    df = load_latest_data()
    if df is None:
        return [], ""
    
    try:
        # 初始化圖表生成器
        chart_generator = ChartGenerator(output_dir="real_charts")
        
        # 生成圖表
        chart_paths = chart_generator.generate_all_charts(df)
        summary_text = chart_generator.generate_summary_text(df)
        
        logger.info(f"✅ 已生成 {len(chart_paths)} 個圖表:")
        for path in chart_paths:
            logger.info(f"  📈 {os.path.basename(path)}")
        
        logger.info("📝 摘要文字:")
        print(summary_text)
        
        return chart_paths, summary_text
        
    except Exception as e:
        logger.error(f"❌ 圖表生成失敗: {e}")
        return [], ""

def send_real_charts_to_telegram(chart_paths, summary_text):
    """發送真實圖表到Telegram"""
    logger.info("📱 發送真實圖表到Telegram...")
    
    try:
        # 初始化Telegram通知器
        notifier = TelegramNotifier()
        
        if not notifier.is_configured():
            logger.warning("⚠️ Telegram未配置，無法發送")
            return False
        
        # 測試連線
        if not notifier.test_connection():
            logger.error("❌ Telegram連線失敗")
            return False
        
        # 準備報告訊息
        report_message = f"""
🤖 **台期所每日自動報告**

{summary_text}

📊 本報告基於台期所公開三大法人資料
⏰ 自動生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        """
        
        # 發送報告
        success = notifier.send_chart_report(chart_paths, report_message)
        
        if success:
            logger.info("🎉 真實圖表報告發送成功！")
            return True
        else:
            logger.warning("⚠️ 圖表報告發送部分失敗")
            return False
            
    except Exception as e:
        logger.error(f"❌ Telegram發送失敗: {e}")
        return False

def main():
    """主程式"""
    logger.info("=== 台期所真實資料Telegram圖表發送測試 ===\n")
    
    # 1. 生成圖表
    chart_paths, summary_text = generate_charts_with_real_data()
    
    if not chart_paths:
        logger.error("❌ 圖表生成失敗，無法繼續測試")
        return 1
    
    # 2. 發送到Telegram
    telegram_success = send_real_charts_to_telegram(chart_paths, summary_text)
    
    # 3. 總結
    logger.info("\n=== 測試結果摘要 ===")
    logger.info(f"📊 圖表生成: {'✅ 成功' if chart_paths else '❌ 失敗'}")
    logger.info(f"📱 Telegram發送: {'✅ 成功' if telegram_success else '❌ 失敗'}")
    
    if chart_paths and telegram_success:
        logger.info("🎉 真實資料圖表發送測試完全成功！")
        logger.info("💡 每日workflow將會自動執行類似的功能")
        return 0
    else:
        logger.error("❌ 部分功能測試失敗，請檢查配置")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 