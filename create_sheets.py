#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
建立Google試算表
"""

from google_sheets_manager import GoogleSheetsManager
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(message)s')

def main():
    print("🚀 正在建立Google試算表...")
    
    # 建立管理器
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        print("請確認認證檔案是否正確設定")
        return
    
    # 建立新的試算表
    spreadsheet = manager.create_spreadsheet('台期所資料分析')
    
    if spreadsheet:
        print(f"\n🎉 試算表建立成功！")
        print(f"📊 試算表名稱: {spreadsheet.title}")
        print(f"🆔 試算表ID: {spreadsheet.id}")
        print(f"🌐 試算表網址: {manager.get_spreadsheet_url()}")
        
        # 設定為公開可檢視
        print("\n🌍 正在設定分享權限...")
        if manager.share_spreadsheet():
            print("✅ 試算表已設定為公開可檢視")
            print("   任何有連結的人都能查看資料")
        
        print("\n📱 手機存取方式:")
        print("1. 下載「Google試算表」APP")
        print("2. 登入你的Google帳號")
        print("3. 找到「台期所資料分析」試算表")
        
        print("\n💻 電腦存取方式:")
        print("1. 開啟瀏覽器")
        print("2. 貼上上面的網址")
        print("3. 即可查看資料")
        
        print(f"\n🔖 請記住這個網址:")
        print(f"🌐 {manager.get_spreadsheet_url()}")
        
    else:
        print("❌ 試算表建立失敗")
        print("請檢查Google Cloud設定是否正確")

if __name__ == "__main__":
    main() 