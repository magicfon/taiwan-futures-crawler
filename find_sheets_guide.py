#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Sheets 位置指南
幫助您找到和存取所有相關的試算表
"""

from google_sheets_manager import GoogleSheetsManager
import logging

def show_sheets_locations():
    """顯示所有Google Sheets的位置和存取方式"""
    
    print("🌐 === Google Sheets 位置指南 ===\n")
    
    print("📍 **主要試算表位置**:")
    print("├── 試算表名稱: 台期所資料分析_新版")
    print("├── 直接連結: https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4")
    print("├── 擁有者: sheets-automation@magicfon.iam.gserviceaccount.com")
    print("└── 共享給: magicfon@gmail.com (編輯權限)\n")
    
    print("🔍 **如何在不同平台找到試算表**:\n")
    
    print("💻 **電腦瀏覽器**:")
    print("1. 前往 https://drive.google.com")
    print("2. 登入您的Gmail帳戶 (magicfon@gmail.com)")
    print("3. 在左側選單點擊「與我共用」")
    print("4. 找到「台期所資料分析_新版」")
    print("5. 或直接開啟連結: https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4\n")
    
    print("📱 **手機 Android/iPhone**:")
    print("1. 下載「Google 試算表」APP")
    print("2. 登入您的Gmail帳戶")
    print("3. 在主畫面會看到「台期所資料分析_新版」")
    print("4. 或在「與我共用」分頁中找到\n")
    
    print("🖥️ **Google Drive APP**:")
    print("1. 下載「Google 雲端硬碟」APP")
    print("2. 登入後點擊「與我共用」")
    print("3. 找到「台期所資料分析_新版」\n")
    
    print("📂 **工作表結構**:")
    print("├── 📊 歷史資料 - 所有爬取的期貨資料")
    print("├── 📈 每日摘要 - 自動生成的每日報告") 
    print("├── 📉 三大法人趨勢 - 趨勢分析資料")
    print("└── ⚙️ 系統資訊 - 系統狀態和說明\n")
    
    print("💾 **與本地檔案的關係**:")
    print("├── 本地備份: output/ 和 output_history/ 目錄")
    print("├── 自動同步: 每次爬取後自動上傳到Google Sheets")
    print("├── 即時存取: Google Sheets隨時可用，不受電腦開關機影響")
    print("└── 離線功能: Google Sheets APP支援離線檢視\n")

def check_current_sheets():
    """檢查目前可存取的試算表"""
    print("🔍 **檢查目前的Google Sheets連接**:\n")
    
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        return
    
    # 連接試算表
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
    spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
    
    if spreadsheet:
        print(f"✅ 成功連接試算表: {spreadsheet.title}")
        print(f"📊 試算表ID: {spreadsheet.id}")
        print(f"🌐 直接連結: {spreadsheet_url}")
        
        # 列出所有工作表
        worksheets = spreadsheet.worksheets()
        print(f"\n📋 包含的工作表:")
        for i, ws in enumerate(worksheets, 1):
            print(f"  {i}. {ws.title}")
        
        # 檢查最近更新時間
        try:
            system_ws = spreadsheet.worksheet("系統資訊")
            last_update = system_ws.acell('B1').value
            print(f"\n⏰ 最後更新時間: {last_update}")
        except:
            print(f"\n⏰ 無法取得最後更新時間")
            
    else:
        print("❌ 無法連接試算表")

def create_bookmark_file():
    """建立書籤檔案，方便快速存取"""
    bookmark_content = """
# 🔖 台期所資料分析 - 快速存取

## 主要連結
- [台期所資料分析_新版](https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4)

## 分頁直接連結
- [歷史資料](https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4/edit#gid=0)
- [每日摘要](https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4/edit#gid=1)
- [三大法人趨勢](https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4/edit#gid=2)
- [系統資訊](https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4/edit#gid=3)

## 快速存取方式
1. **桌面捷徑**: 將主要連結加到瀏覽器書籤
2. **手機APP**: 安裝Google試算表APP
3. **通知設定**: 在Google Sheets中設定變更通知

## 本地備份位置
- 輸出目錄: `output/` 和 `output_history/`
- 自動備份: 每次爬取都會產生本地CSV檔案
"""
    
    with open("GOOGLE_SHEETS_BOOKMARKS.md", "w", encoding="utf-8") as f:
        f.write(bookmark_content)
    
    print("📋 已建立書籤檔案: GOOGLE_SHEETS_BOOKMARKS.md")

def main():
    """主程式"""
    show_sheets_locations()
    check_current_sheets()
    
    print(f"\n" + "="*60)
    
    create_bookmark = input("是否建立書籤檔案？(y/n): ").strip().lower()
    if create_bookmark in ['y', 'yes']:
        create_bookmark_file()
    
    print(f"\n🎯 **重點提醒**:")
    print(f"✅ Google Sheets存在雲端，任何時間任何地點都能存取")
    print(f"✅ 本地也有備份在 output/ 目錄")
    print(f"✅ 手機APP可以離線檢視資料")
    print(f"✅ 支援即時協作和分享")

if __name__ == "__main__":
    main() 