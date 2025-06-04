#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試每日爬取和重試機制
模擬GitHub Actions的行為
"""

import subprocess
import time
import os
import logging
from datetime import datetime

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("測試爬蟲")

def check_output_files():
    """檢查是否有產生輸出檔案"""
    output_dir = "output"
    if not os.path.exists(output_dir):
        return False
    
    files = os.listdir(output_dir)
    # 過濾出今天產生的檔案
    today_str = datetime.now().strftime('%Y%m%d')
    today_files = [f for f in files if today_str in f and (f.endswith('.csv') or f.endswith('.xlsx'))]
    
    return len(today_files) > 0

def run_crawler():
    """執行爬蟲並回傳是否成功"""
    try:
        logger.info("🚀 開始執行爬蟲...")
        
        # 執行今日資料爬取
        result = subprocess.run([
            'python', 'taifex_crawler.py', 
            '--date-range', 'today',
            '--contracts', 'TX,TE,MTX'
        ], capture_output=True, text=True, timeout=300)
        
        logger.info(f"爬蟲退出碼: {result.returncode}")
        
        if result.stdout:
            logger.info("標準輸出:")
            for line in result.stdout.split('\n')[-10:]:  # 只顯示最後10行
                if line.strip():
                    logger.info(f"  {line}")
        
        if result.stderr:
            logger.warning("錯誤輸出:")
            for line in result.stderr.split('\n')[-5:]:  # 只顯示最後5行
                if line.strip():
                    logger.warning(f"  {line}")
        
        # 檢查是否有產生檔案
        has_files = check_output_files()
        
        # 爬蟲成功條件：退出碼為0且有產生檔案
        success = (result.returncode == 0) and has_files
        
        if success:
            logger.info("✅ 爬蟲執行成功，有產生資料檔案")
        else:
            if result.returncode != 0:
                logger.warning(f"⚠️ 爬蟲退出碼異常: {result.returncode}")
            if not has_files:
                logger.warning("⚠️ 沒有產生資料檔案")
        
        return success
        
    except subprocess.TimeoutExpired:
        logger.error("❌ 爬蟲執行超時")
        return False
    except Exception as e:
        logger.error(f"❌ 爬蟲執行失敗: {e}")
        return False

def main():
    """主測試程式"""
    logger.info("=== 測試每日爬取和重試機制 ===")
    
    # 第一次嘗試
    logger.info("\n📋 第一次嘗試:")
    first_success = run_crawler()
    
    if first_success:
        logger.info("🎉 第一次嘗試成功，無需重試")
        return 0
    
    # 如果第一次失敗，等待後重試
    logger.info("\n⏰ 第一次失敗，等待10秒後重試...")
    time.sleep(10)  # 實際環境是600秒(10分鐘)，測試環境縮短為10秒
    
    logger.info("\n🔄 重試嘗試:")
    retry_success = run_crawler()
    
    if retry_success:
        logger.info("🎉 重試成功！")
        return 0
    else:
        logger.error("❌ 兩次嘗試都失敗")
        return 1

def show_current_time_info():
    """顯示當前時間資訊"""
    now = datetime.now()
    logger.info(f"當前時間: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"星期: {['週一', '週二', '週三', '週四', '週五', '週六', '週日'][now.weekday()]}")
    
    if now.weekday() >= 5:
        logger.info("今天是週末，可能沒有交易資料")
    else:
        logger.info("今天是工作日，應該有交易資料")

if __name__ == "__main__":
    show_current_time_info()
    exit_code = main()
    print(f"\n最終退出碼: {exit_code}")
    exit(exit_code) 