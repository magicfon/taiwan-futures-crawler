#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
分享Google Sheets給個人帳戶
解決服務帳戶建立的試算表權限問題
"""

from google_sheets_manager import GoogleSheetsManager
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def share_sheets():
    """分享試算表給個人帳戶或設定為公開"""
    print("🔐 正在設定Google Sheets分享權限...")
    
    # 連接到現有試算表
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗，請檢查認證設定")
        return False
    
    # 連接到我們的試算表
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
    spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
    
    if not spreadsheet:
        print("❌ 連接試算表失敗")
        return False
    
    print(f"✅ 成功連接試算表: {spreadsheet.title}")
    
    # 提供選項
    print(f"\n📋 分享選項:")
    print(f"1. 設定為公開可檢視 (任何人都可以查看)")
    print(f"2. 分享給特定Email帳戶")
    print(f"3. 取消")
    
    choice = input(f"\n請選擇 (1/2/3): ").strip()
    
    if choice == "1":
        # 設定為公開
        try:
            success = manager.share_spreadsheet()
            if success:
                print(f"✅ 試算表已設定為公開可檢視!")
                print(f"🌐 任何人都可以透過此連結查看: {spreadsheet_url}")
                return True
            else:
                print(f"❌ 設定公開權限失敗")
                return False
        except Exception as e:
            print(f"❌ 設定公開權限失敗: {e}")
            return False
    
    elif choice == "2":
        # 分享給特定Email
        email = input(f"請輸入要分享的Email地址: ").strip()
        if not email:
            print(f"❌ Email地址不能為空")
            return False
        
        role_choice = input(f"權限類型 (1=檢視者/2=編輯者) [預設:1]: ").strip()
        role = 'writer' if role_choice == "2" else 'reader'
        
        try:
            success = manager.share_spreadsheet(email, role)
            if success:
                role_text = "編輯者" if role == 'writer' else "檢視者"
                print(f"✅ 試算表已分享給 {email} ({role_text})!")
                print(f"🌐 連結: {spreadsheet_url}")
                return True
            else:
                print(f"❌ 分享失敗")
                return False
        except Exception as e:
            print(f"❌ 分享失敗: {e}")
            return False
    
    elif choice == "3":
        print(f"🚫 已取消")
        return False
    
    else:
        print(f"❌ 無效選擇")
        return False

def main():
    """主程式"""
    print("=== Google Sheets 分享工具 ===\n")
    
    success = share_sheets()
    
    if success:
        print(f"\n🎉 分享設定完成!")
        print(f"📋 現在您可以:")
        print(f"1. 在瀏覽器中開啟試算表連結")
        print(f"2. 在手機上使用Google試算表APP")
        print(f"3. 分享連結給其他需要的人")
    else:
        print(f"\n😞 分享設定失敗")
        print(f"📋 請檢查:")
        print(f"1. 網路連接是否正常")
        print(f"2. Google認證是否有效")
        print(f"3. Email地址是否正確")

if __name__ == "__main__":
    main() 