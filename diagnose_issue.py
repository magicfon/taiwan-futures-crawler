#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
診斷腳本 - 檢查為什麼Google Sheets和Telegram功能沒有執行
"""

import os
import sys
import datetime
from pathlib import Path

def check_environment():
    """檢查環境設定"""
    print("🔍 檢查環境設定...")
    
    # 檢查GitHub Secrets環境變數
    google_cred = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
    telegram_token = os.getenv('TELEGRAM_BOT_TOKEN')
    telegram_chat = os.getenv('TELEGRAM_CHAT_ID')
    
    print(f"GOOGLE_SHEETS_CREDENTIALS: {'✅ 已設定' if google_cred else '❌ 未設定'}")
    print(f"TELEGRAM_BOT_TOKEN: {'✅ 已設定' if telegram_token else '❌ 未設定'}")
    print(f"TELEGRAM_CHAT_ID: {'✅ 已設定' if telegram_chat else '❌ 未設定'}")
    
    # 檢查本地認證檔案
    local_google_cred = Path("config/google_sheets_credentials.json")
    print(f"本地Google認證檔案: {'✅ 存在' if local_google_cred.exists() else '❌ 不存在'}")
    
    return google_cred or local_google_cred.exists(), telegram_token and telegram_chat

def check_database():
    """檢查資料庫狀況"""
    print("🔍 檢查資料庫狀況...")
    
    try:
        from database_manager import TaifexDatabaseManager
        
        db_manager = TaifexDatabaseManager()
        
        # 檢查最近30天資料
        recent_data = db_manager.get_recent_data(30)
        print(f"最近30天資料筆數: {len(recent_data)}")
        
        if not recent_data.empty:
            print(f"最新資料日期: {recent_data['date'].max()}")
            print(f"最舊資料日期: {recent_data['date'].min()}")
            
            # 檢查是否超過觸發門檻
            if len(recent_data) > 50:
                print("✅ 資料量足夠觸發Google Sheets和Telegram功能")
                return True, recent_data
            else:
                print(f"⚠️ 資料量不足（{len(recent_data)}/50），需要更多歷史資料")
                return False, recent_data
        else:
            print("❌ 資料庫中沒有最近30天的資料")
            return False, recent_data
            
    except Exception as e:
        print(f"❌ 資料庫檢查失敗: {e}")
        return False, None

def check_modules():
    """檢查相關模組是否可用"""
    print("🔍 檢查模組可用性...")
    
    modules = {
        'Google Sheets': 'google_sheets_manager',
        'Telegram': 'telegram_notifier', 
        'Chart Generation': 'chart_generator',
        'Database': 'database_manager'
    }
    
    available = {}
    
    for name, module in modules.items():
        try:
            __import__(module)
            print(f"{name}: ✅ 可用")
            available[name] = True
        except ImportError as e:
            print(f"{name}: ❌ 不可用 ({e})")
            available[name] = False
    
    return available

def simulate_workflow():
    """模擬工作流程，找出問題點"""
    print("🔍 模擬工作流程...")
    
    # 1. 檢查環境
    google_ok, telegram_ok = check_environment()
    
    # 2. 檢查模組
    modules = check_modules()
    
    # 3. 檢查資料庫
    db_ok, recent_data = check_database()
    
    print("\n" + "="*50)
    print("📋 診斷結果")
    print("="*50)
    
    # 分析為什麼功能沒有執行
    issues = []
    
    if not google_ok:
        issues.append("❌ Google Sheets認證未正確設定")
    
    if not telegram_ok:
        issues.append("❌ Telegram認證未正確設定")
    
    if not modules.get('Google Sheets', False):
        issues.append("❌ Google Sheets模組不可用")
    
    if not modules.get('Telegram', False):
        issues.append("❌ Telegram模組不可用")
        
    if not modules.get('Chart Generation', False):
        issues.append("❌ 圖表生成模組不可用")
    
    if not db_ok:
        issues.append("❌ 資料庫資料不足或無資料")
    
    if issues:
        print("發現以下問題:")
        for issue in issues:
            print(f"  {issue}")
    else:
        print("✅ 所有條件都滿足，功能應該正常執行")
    
    print("\n💡 解決建議:")
    
    if not google_ok:
        print("- 確認已在GitHub Secrets中設定 GOOGLE_SHEETS_CREDENTIALS")
        print("- 或在config/目錄下放置 google_sheets_credentials.json")
    
    if not telegram_ok:
        print("- 確認已在GitHub Secrets中設定 TELEGRAM_BOT_TOKEN 和 TELEGRAM_CHAT_ID")
    
    if not db_ok:
        print("- 需要先成功爬取一些歷史資料")
        print("- 嘗試執行: python taifex_crawler.py --date-range 2024-11-01,2024-12-06")
    
    return len(issues) == 0

def check_recent_crawl():
    """檢查最近的爬蟲執行狀況"""
    print("🔍 檢查最近的爬蟲執行...")
    
    # 檢查輸出目錄
    output_dir = Path("output")
    if output_dir.exists():
        files = list(output_dir.glob("*"))
        if files:
            print(f"輸出目錄包含 {len(files)} 個檔案:")
            for f in sorted(files, key=lambda x: x.stat().st_mtime, reverse=True)[:5]:
                mtime = datetime.datetime.fromtimestamp(f.stat().st_mtime)
                print(f"  📄 {f.name} (修改時間: {mtime.strftime('%Y-%m-%d %H:%M:%S')})")
        else:
            print("❌ 輸出目錄為空")
    else:
        print("❌ 輸出目錄不存在")
    
    # 檢查資料庫檔案
    db_file = Path("data/taifex_data.db")
    if db_file.exists():
        mtime = datetime.datetime.fromtimestamp(db_file.stat().st_mtime)
        print(f"✅ 資料庫檔案存在 (修改時間: {mtime.strftime('%Y-%m-%d %H:%M:%S')})")
    else:
        print("❌ 資料庫檔案不存在")

def main():
    """主診斷函數"""
    print("=" * 60)
    print("🩺 台期所爬蟲功能診斷工具")
    print("=" * 60)
    print(f"診斷時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 1. 檢查最近執行狀況
    check_recent_crawl()
    print()
    
    # 2. 模擬工作流程
    all_ok = simulate_workflow()
    
    if all_ok:
        print("\n🎉 診斷完成：所有條件都滿足，功能應該正常執行")
        print("💡 如果功能仍未執行，請檢查GitHub Actions的執行日誌")
        return 0
    else:
        print("\n⚠️ 診斷完成：發現一些需要解決的問題")
        print("📋 請按照上述建議修正問題後重新嘗試")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code) 