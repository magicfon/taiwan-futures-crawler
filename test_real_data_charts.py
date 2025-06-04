#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試真實資料的圖表生成和Telegram發送
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging
import os
import glob

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("真實資料圖表測試")

def load_real_data():
    """從CSV檔案載入真實的爬蟲資料"""
    logger.info("🔍 搜尋CSV檔案...")
    
    # 搜尋最新的CSV檔案
    csv_files = glob.glob("output/*.csv")
    if not csv_files:
        logger.error("❌ 在output目錄中沒有找到CSV檔案")
        return pd.DataFrame()
    
    # 選取最新的檔案
    latest_csv = max(csv_files, key=os.path.getmtime)
    logger.info(f"📄 載入檔案: {latest_csv}")
    
    try:
        df = pd.read_csv(latest_csv, encoding='utf-8-sig')
        logger.info(f"✅ 成功載入 {len(df)} 筆資料")
        
        # 顯示資料結構
        logger.info(f"📊 資料欄位: {df.columns.tolist()}")
        if not df.empty:
            logger.info(f"📅 日期範圍: {df.iloc[0].get('日期', 'N/A')} 到 {df.iloc[-1].get('日期', 'N/A')}")
            logger.info(f"📈 契約類型: {df['契約名稱'].unique().tolist() if '契約名稱' in df.columns else 'N/A'}")
        
        return df
        
    except Exception as e:
        logger.error(f"❌ 讀取CSV檔案失敗: {e}")
        return pd.DataFrame()

def convert_data_format(df):
    """將CSV資料轉換為圖表生成器需要的格式"""
    if df.empty:
        return df
    
    # 檢查必要欄位是否存在
    required_columns = ['日期', '契約名稱']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        logger.error(f"❌ 缺少必要欄位: {missing_columns}")
        return pd.DataFrame()
    
    # 轉換日期格式
    df['日期'] = pd.to_datetime(df['日期'])
    
    # 確保有多空淨額欄位
    if '多空淨額交易口數' not in df.columns:
        df['多空淨額交易口數'] = 0
    if '多空淨額未平倉口數' not in df.columns:
        df['多空淨額未平倉口數'] = 0
    
    # 填充NaN值
    df['多空淨額交易口數'] = df['多空淨額交易口數'].fillna(0)
    df['多空淨額未平倉口數'] = df['多空淨額未平倉口數'].fillna(0)
    
    logger.info("✅ 資料格式轉換完成")
    return df

def generate_sample_data():
    """如果沒有真實資料，生成範例資料"""
    logger.info("🔧 生成範例資料...")
    
    # 生成30天的範例資料
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
                # 生成模擬的真實感資料
                base_volume = {
                    'TX': np.random.randint(-8000, 8000),
                    'TE': np.random.randint(-3000, 3000),
                    'MTX': np.random.randint(-2000, 2000)
                }
                
                base_position = {
                    'TX': np.random.randint(-15000, 15000),
                    'TE': np.random.randint(-8000, 8000),
                    'MTX': np.random.randint(-5000, 5000)
                }
                
                test_data.append({
                    '日期': date,
                    '契約名稱': contract,
                    '身份別': identity,
                    '多空淨額交易口數': base_volume[contract],
                    '多空淨額未平倉口數': base_position[contract]
                })
    
    df = pd.DataFrame(test_data)
    logger.info(f"✅ 已生成 {len(df)} 筆範例資料")
    return df

def test_chart_and_telegram():
    """測試圖表生成和Telegram發送"""
    try:
        from chart_generator import ChartGenerator
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== 真實資料圖表生成和Telegram發送測試 ===\n")
        
        # 1. 載入資料
        df = load_real_data()
        
        # 如果沒有真實資料，使用範例資料
        if df.empty:
            logger.warning("⚠️ 沒有找到真實資料，使用範例資料")
            df = generate_sample_data()
        
        # 2. 轉換資料格式
        df = convert_data_format(df)
        
        if df.empty:
            logger.error("❌ 沒有有效資料可生成圖表")
            return False
        
        # 3. 生成圖表
        logger.info("🎨 開始生成圖表...")
        chart_generator = ChartGenerator(output_dir="real_charts")
        
        chart_paths = chart_generator.generate_all_charts(df)
        summary_text = chart_generator.generate_summary_text(df)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return False
        
        logger.info(f"✅ 已生成 {len(chart_paths)} 個圖表:")
        for path in chart_paths:
            logger.info(f"  📈 {os.path.basename(path)}")
        
        # 4. 發送到Telegram
        logger.info("📱 開始發送到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        # 測試連線
        if not notifier.test_connection():
            logger.error("❌ Telegram連線失敗")
            return False
        
        # 發送報告
        success = notifier.send_chart_report(chart_paths, summary_text)
        
        if success:
            logger.info("🎉 圖表已成功發送到Telegram！")
            logger.info("📱 請檢查您的Telegram，應該已收到台期所持倉分析圖表")
            return True
        else:
            logger.error("❌ Telegram發送失敗")
            return False
            
    except ImportError as e:
        logger.error(f"❌ 模組載入失敗: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ 執行過程中發生錯誤: {e}")
        return False

def main():
    """主程式"""
    success = test_chart_and_telegram()
    
    if success:
        logger.info("\n🎊 測試完成！")
        logger.info("📊 台期所資料圖表已成功生成並發送到Telegram")
        logger.info("💡 現在您可以設定每日自動執行，定期收到最新的持倉分析了")
        return 0
    else:
        logger.error("\n❌ 測試失敗，請檢查錯誤訊息")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 