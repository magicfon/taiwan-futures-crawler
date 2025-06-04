#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
建立新的Google Sheets資料庫
由您的服務帳戶完全控制，避免權限問題
"""

from google_sheets_manager import GoogleSheetsManager
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_new_database():
    """建立新的台期所資料庫"""
    print("🚀 正在建立新的台期所Google Sheets資料庫...")
    
    # 初始化管理器
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗，請檢查認證設定")
        return None
    
    # 建立新試算表
    spreadsheet = manager.create_spreadsheet("台期所資料分析_新版")
    
    if spreadsheet:
        print(f"✅ 成功建立新的資料庫!")
        print(f"📊 試算表名稱: {spreadsheet.title}")
        print(f"🌐 URL: {manager.get_spreadsheet_url()}")
        print(f"🔑 試算表ID: {spreadsheet.id}")
        
        # 更新系統資訊
        manager.update_system_info()
        
        # 提供更新指示
        print(f"\n📋 請更新以下檔案:")
        print(f"1. crawl_history.py - 將舊URL替換為新URL")
        print(f"2. 任何其他使用舊URL的腳本")
        print(f"\n新URL: {manager.get_spreadsheet_url()}")
        
        return spreadsheet
    else:
        print("❌ 建立資料庫失敗")
        return None

def update_crawl_history_url(new_url):
    """更新crawl_history.py中的URL"""
    try:
        with open('crawl_history.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 替換舊URL
        old_url = "https://docs.google.com/spreadsheets/d/1Ltv8zsQcCQ5SiaYKsgCDetNC-SqEMZP4V33S2nKMuWI"
        new_content = content.replace(old_url, new_url)
        
        with open('crawl_history.py', 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        print(f"✅ 已更新 crawl_history.py 中的URL")
        return True
    except Exception as e:
        print(f"❌ 更新檔案失敗: {e}")
        return False

def main():
    """主程式"""
    print("=== 台期所Google Sheets資料庫重建工具 ===\n")
    
    # 建立新資料庫
    spreadsheet = create_new_database()
    
    if spreadsheet:
        new_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
        
        # 詢問是否自動更新程式碼
        user_input = input(f"\n是否自動更新 crawl_history.py 中的URL？(y/n): ")
        if user_input.lower() in ['y', 'yes', 'Y']:
            update_crawl_history_url(new_url)
        
        print(f"\n🎉 設定完成!")
        print(f"📋 下一步:")
        print(f"1. 執行 python crawl_history.py 來測試新資料庫")
        print(f"2. 檢查資料是否正確上傳")
        print(f"3. 分享試算表給需要的使用者")

if __name__ == "__main__":
    main() 