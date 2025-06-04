#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
連接到現有的Google試算表
"""

from google_sheets_manager import GoogleSheetsManager
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(message)s')

def main():
    # 你的現有試算表網址
    existing_url = "https://docs.google.com/spreadsheets/d/1ibPtmvy2rZN8Lke1BOnlxq1udMmFTn3Xg1gY3vdlx8s/edit?gid=1710574419#gid=1710574419"
    
    print("🔗 正在連接到你的現有Google試算表...")
    print(f"📊 試算表網址: {existing_url}")
    
    # 建立管理器
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        print("還是需要Google Cloud Console設定來取得API認證")
        print("這是Google的安全規定，程式需要「鑰匙」才能自動上傳資料")
        return
    
    # 連接到現有試算表
    spreadsheet = manager.connect_spreadsheet(existing_url)
    
    if spreadsheet:
        print(f"\n🎉 成功連接到現有試算表！")
        print(f"📊 試算表名稱: {spreadsheet.title}")
        print(f"🆔 試算表ID: {spreadsheet.id}")
        print(f"🌐 試算表網址: {manager.get_spreadsheet_url()}")
        
        print(f"\n📋 現有工作表:")
        for i, worksheet in enumerate(spreadsheet.worksheets(), 1):
            print(f"   {i}. {worksheet.title}")
        
        print(f"\n✅ 現在程式可以自動上傳資料到這個試算表了！")
        print(f"🤖 每日09:00會自動更新資料")
        
        # 測試寫入權限
        print(f"\n🧪 測試寫入權限...")
        try:
            # 嘗試寫入系統資訊
            if manager.update_system_info():
                print("✅ 寫入權限正常，可以自動上傳資料！")
            else:
                print("⚠️ 寫入測試失敗，請檢查權限設定")
        except Exception as e:
            print(f"⚠️ 寫入測試失敗: {e}")
            print("可能需要分享試算表給服務帳號")
            
            # 顯示服務帳號Email
            try:
                with open("config/google_sheets_credentials.json", 'r') as f:
                    import json
                    creds = json.load(f)
                    service_email = creds.get('client_email', '未知')
                    print(f"\n📧 請將試算表分享給: {service_email}")
                    print("分享權限: 編輯者")
            except:
                print("無法讀取服務帳號資訊")
        
    else:
        print("❌ 連接現有試算表失敗")
        print("請檢查網址是否正確，或是否有存取權限")

if __name__ == "__main__":
    main() 