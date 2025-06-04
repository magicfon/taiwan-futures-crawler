#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Telegram 通知模組
用於發送台期所資料分析圖表到 Telegram
"""

import requests
import logging
import os
from pathlib import Path
import time

logger = logging.getLogger("Telegram通知")

class TelegramNotifier:
    def __init__(self, bot_token, chat_id):
        """
        初始化Telegram通知器
        
        Args:
            bot_token: Telegram Bot的API Token
            chat_id: 目標聊天室ID
        """
        self.bot_token = bot_token
        self.chat_id = chat_id
        self.base_url = f"https://api.telegram.org/bot{bot_token}"
        
    def send_message(self, text, parse_mode="Markdown"):
        """
        發送文字訊息
        
        Args:
            text: 要發送的文字內容
            parse_mode: 解析模式 (Markdown 或 HTML)
            
        Returns:
            bool: 是否發送成功
        """
        try:
            url = f"{self.base_url}/sendMessage"
            payload = {
                'chat_id': self.chat_id,
                'text': text,
                'parse_mode': parse_mode
            }
            
            response = requests.post(url, data=payload, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    logger.info("📱 Telegram訊息發送成功")
                    return True
                else:
                    logger.error(f"Telegram API錯誤: {result.get('description', '未知錯誤')}")
                    return False
            else:
                logger.error(f"HTTP錯誤: {response.status_code}")
                return False
                
        except requests.exceptions.RequestException as e:
            logger.error(f"網路請求失敗: {e}")
            return False
        except Exception as e:
            logger.error(f"發送訊息時發生錯誤: {e}")
            return False
    
    def send_photo(self, photo_path, caption="", parse_mode="Markdown"):
        """
        發送圖片
        
        Args:
            photo_path: 圖片檔案路徑
            caption: 圖片說明文字
            parse_mode: 解析模式
            
        Returns:
            bool: 是否發送成功
        """
        try:
            if not os.path.exists(photo_path):
                logger.error(f"圖片檔案不存在: {photo_path}")
                return False
            
            url = f"{self.base_url}/sendPhoto"
            
            with open(photo_path, 'rb') as photo_file:
                files = {'photo': photo_file}
                data = {
                    'chat_id': self.chat_id,
                    'caption': caption,
                    'parse_mode': parse_mode
                }
                
                response = requests.post(url, files=files, data=data, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    logger.info(f"📸 圖片發送成功: {os.path.basename(photo_path)}")
                    return True
                else:
                    logger.error(f"Telegram API錯誤: {result.get('description', '未知錯誤')}")
                    return False
            else:
                logger.error(f"HTTP錯誤: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"發送圖片時發生錯誤: {e}")
            return False
    
    def send_document(self, document_path, caption="", parse_mode="Markdown"):
        """
        發送文件
        
        Args:
            document_path: 文件檔案路徑
            caption: 文件說明文字
            parse_mode: 解析模式
            
        Returns:
            bool: 是否發送成功
        """
        try:
            if not os.path.exists(document_path):
                logger.error(f"文件不存在: {document_path}")
                return False
            
            url = f"{self.base_url}/sendDocument"
            
            with open(document_path, 'rb') as doc_file:
                files = {'document': doc_file}
                data = {
                    'chat_id': self.chat_id,
                    'caption': caption,
                    'parse_mode': parse_mode
                }
                
                response = requests.post(url, files=files, data=data, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    logger.info(f"📄 文件發送成功: {os.path.basename(document_path)}")
                    return True
                else:
                    logger.error(f"Telegram API錯誤: {result.get('description', '未知錯誤')}")
                    return False
            else:
                logger.error(f"HTTP錯誤: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"發送文件時發生錯誤: {e}")
            return False
    
    def send_chart_report(self, chart_paths, summary_text=""):
        """
        發送圖表報告
        
        Args:
            chart_paths: 圖表檔案路徑列表
            summary_text: 摘要文字
            
        Returns:
            bool: 是否全部發送成功
        """
        success_count = 0
        total_count = len(chart_paths)
        
        # 發送摘要訊息
        if summary_text:
            if self.send_message(summary_text):
                success_count += 1
        
        # 發送每個圖表
        for i, chart_path in enumerate(chart_paths):
            if os.path.exists(chart_path):
                # 獲取契約名稱作為說明
                filename = os.path.basename(chart_path)
                contract_name = filename.split('_')[0] if '_' in filename else "契約分析"
                
                caption = f"📊 {contract_name} 持倉分析圖表\n⏰ 更新時間: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                
                if self.send_photo(chart_path, caption):
                    success_count += 1
                    
                # 避免發送過快被限制
                if i < len(chart_paths) - 1:
                    time.sleep(1)
            else:
                logger.warning(f"圖表檔案不存在: {chart_path}")
        
        logger.info(f"Telegram發送完成: {success_count}/{total_count + (1 if summary_text else 0)} 成功")
        return success_count == total_count + (1 if summary_text else 0)
    
    def test_connection(self):
        """
        測試Telegram連線
        
        Returns:
            bool: 連線是否正常
        """
        try:
            url = f"{self.base_url}/getMe"
            response = requests.get(url, timeout=10)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    bot_info = result.get('result', {})
                    logger.info(f"✅ Telegram Bot連線正常: {bot_info.get('first_name', 'Unknown')}")
                    return True
                else:
                    logger.error(f"Telegram API錯誤: {result.get('description', '未知錯誤')}")
                    return False
            else:
                logger.error(f"HTTP錯誤: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"測試連線時發生錯誤: {e}")
            return False

# 測試函數
def test_telegram_notifier():
    """測試Telegram通知功能"""
    # 使用用戶提供的認證資訊
    bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
    chat_id = "1038401606"
    
    notifier = TelegramNotifier(bot_token, chat_id)
    
    # 測試連線
    if notifier.test_connection():
        # 發送測試訊息
        test_message = "🤖 台期所資料爬蟲系統測試訊息\n📊 系統運作正常"
        notifier.send_message(test_message)
    else:
        logger.error("Telegram連線測試失敗")

if __name__ == "__main__":
    test_telegram_notifier() 