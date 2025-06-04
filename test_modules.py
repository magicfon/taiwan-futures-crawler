#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試模組導入狀況
"""

print("🧪 測試模組導入...")

# 測試Google Sheets模組
try:
    from google_sheets_manager import GoogleSheetsManager
    print("✅ Google Sheets模組導入成功")
    SHEETS_AVAILABLE = True
except Exception as e:
    print(f"❌ Google Sheets模組導入失敗: {e}")
    SHEETS_AVAILABLE = False

# 測試資料庫模組
try:
    from database_manager import TaifexDatabaseManager
    print("✅ 資料庫模組導入成功")
    DB_AVAILABLE = True
except Exception as e:
    print(f"❌ 資料庫模組導入失敗: {e}")
    DB_AVAILABLE = False

# 測試日報模組
try:
    from daily_report_generator import DailyReportGenerator
    print("✅ 日報模組導入成功")
    REPORT_AVAILABLE = True
except Exception as e:
    print(f"❌ 日報模組導入失敗: {e}")
    REPORT_AVAILABLE = False

print(f"\n📊 模組狀況總結:")
print(f"Google Sheets: {'✅' if SHEETS_AVAILABLE else '❌'}")
print(f"資料庫: {'✅' if DB_AVAILABLE else '❌'}")
print(f"日報: {'✅' if REPORT_AVAILABLE else '❌'}")

if SHEETS_AVAILABLE:
    print(f"\n🔗 測試Google Sheets連接...")
    manager = GoogleSheetsManager()
    if manager.client:
        print("✅ Google Sheets連接成功")
    else:
        print("❌ Google Sheets連接失敗") 