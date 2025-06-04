#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試從Google Sheets載入歷史資料並生成圖表
"""

import pandas as pd
import logging
import os
from datetime import datetime

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("Google Sheets圖表測試")

def test_google_sheets_charts():
    """測試從Google Sheets載入資料並生成圖表"""
    try:
        from chart_generator import ChartGenerator
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== Google Sheets歷史資料圖表生成測試 ===\n")
        
        # 1. 初始化圖表生成器
        chart_generator = ChartGenerator(output_dir="sheets_charts")
        
        # 2. 從Google Sheets載入30天歷史資料
        logger.info("📊 從Google Sheets載入30天歷史資料...")
        chart_data = chart_generator.load_data_from_google_sheets(30)
        
        if chart_data.empty:
            logger.error("❌ 無法從Google Sheets載入資料")
            return False
        
        logger.info(f"✅ 成功載入 {len(chart_data)} 筆歷史資料")
        logger.info(f"📅 日期範圍: {chart_data['日期'].min()} 到 {chart_data['日期'].max()}")
        logger.info(f"📈 契約類型: {chart_data['契約名稱'].unique().tolist() if '契約名稱' in chart_data.columns else 'N/A'}")
        
        # 3. 生成圖表
        logger.info("🎨 開始生成圖表...")
        chart_paths = chart_generator.generate_all_charts(chart_data)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return False
        
        logger.info(f"✅ 已生成 {len(chart_paths)} 個圖表:")
        for path in chart_paths:
            logger.info(f"  📈 {os.path.basename(path)}")
        
        # 4. 生成摘要文字
        summary_text = chart_generator.generate_summary_text(chart_data)
        logger.info("📝 摘要文字:")
        print(summary_text)
        
        # 5. 發送到Telegram
        logger.info("📱 開始發送到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        # 測試連線
        if not notifier.test_connection():
            logger.error("❌ Telegram連線失敗")
            return False
        
        # 發送圖表報告
        success = notifier.send_chart_report(chart_paths, summary_text)
        
        if success:
            logger.info("🎉 完整的30天歷史圖表已成功發送到Telegram！")
            logger.info("📱 請檢查您的Telegram，現在應該顯示完整的30天趨勢分析")
            return True
        else:
            logger.error("❌ Telegram發送失敗")
            return False
            
    except ImportError as e:
        logger.error(f"❌ 模組載入失敗: {e}")
        logger.info("請確保已安裝所有必要套件：pip install -r requirements.txt")
        return False
    except Exception as e:
        logger.error(f"❌ 執行過程中發生錯誤: {e}")
        return False

def show_google_sheets_info():
    """顯示Google Sheets資料資訊"""
    try:
        from google_sheets_manager import GoogleSheetsManager
        import json
        from pathlib import Path
        
        logger.info("📋 Google Sheets資料資訊:")
        
        # 載入配置
        config_file = Path("config/spreadsheet_config.json")
        if config_file.exists():
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                logger.info(f"🔗 試算表URL: {config.get('spreadsheet_url', 'N/A')}")
        
        # 檢查工作表
        sheets_manager = GoogleSheetsManager()
        if sheets_manager.client:
            if config_file.exists():
                spreadsheet_id = config.get('spreadsheet_id')
                if spreadsheet_id:
                    sheets_manager.connect_spreadsheet(spreadsheet_id)
            
            if sheets_manager.spreadsheet:
                worksheets = sheets_manager.spreadsheet.worksheets()
                logger.info("📊 可用工作表:")
                for ws in worksheets:
                    try:
                        row_count = len(ws.get_all_values())
                        logger.info(f"  • {ws.title}: {row_count} 行資料")
                    except:
                        logger.info(f"  • {ws.title}: 無法讀取")
            else:
                logger.warning("❌ 無法連接到Google試算表")
        else:
            logger.warning("❌ Google Sheets未啟用")
            
    except Exception as e:
        logger.error(f"❌ 檢查Google Sheets資訊時發生錯誤: {e}")

def main():
    """主程式"""
    logger.info("=== 開始測試 ===\n")
    
    # 1. 顯示Google Sheets資訊
    show_google_sheets_info()
    
    print("\n" + "="*60 + "\n")
    
    # 2. 測試圖表生成
    success = test_google_sheets_charts()
    
    # 3. 總結
    logger.info("\n=== 測試結果 ===")
    if success:
        logger.info("🎉 測試成功！")
        logger.info("💡 現在圖表應該顯示完整的30天歷史資料")
        logger.info("📱 請檢查Telegram確認圖表內容")
        return 0
    else:
        logger.error("❌ 測試失敗")
        logger.info("💡 請檢查Google Sheets配置和資料")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 