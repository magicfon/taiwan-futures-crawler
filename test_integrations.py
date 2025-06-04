#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
整合功能測試腳本
用於測試Google Sheets和Telegram功能是否正常
"""

import os
import sys
import datetime

def test_google_sheets():
    """測試Google Sheets整合"""
    print("🔍 測試Google Sheets整合...")
    
    try:
        from google_sheets_manager import GoogleSheetsManager
        
        # 檢查認證檔案
        cred_file = "config/google_sheets_credentials.json"
        env_cred = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
        
        if os.path.exists(cred_file):
            print("✅ 找到本地認證檔案")
        elif env_cred:
            print("✅ 找到環境變數認證")
        else:
            print("❌ 沒有找到Google Sheets認證")
            return False
        
        # 嘗試初始化
        sheets_manager = GoogleSheetsManager()
        if sheets_manager.client:
            print("✅ Google Sheets客戶端初始化成功")
            return True
        else:
            print("❌ Google Sheets客戶端初始化失敗")
            return False
            
    except ImportError:
        print("❌ 缺少Google Sheets相關模組")
        return False
    except Exception as e:
        print(f"❌ Google Sheets測試失敗: {e}")
        return False

def test_telegram():
    """測試Telegram整合"""
    print("🔍 測試Telegram整合...")
    
    try:
        from telegram_notifier import TelegramNotifier
        
        # 檢查環境變數
        bot_token = os.getenv('TELEGRAM_BOT_TOKEN')
        chat_id = os.getenv('TELEGRAM_CHAT_ID')
        
        if bot_token and chat_id:
            print("✅ 找到Telegram環境變數")
        else:
            print("❌ 缺少Telegram環境變數")
            return False
        
        # 嘗試初始化
        notifier = TelegramNotifier()
        if notifier.is_configured():
            print("✅ Telegram通知器設定正確")
            
            # 測試連線
            if notifier.test_connection():
                print("✅ Telegram連線測試成功")
                return True
            else:
                print("❌ Telegram連線測試失敗")
                return False
        else:
            print("❌ Telegram通知器設定不完整")
            return False
            
    except ImportError:
        print("❌ 缺少Telegram相關模組")
        return False
    except Exception as e:
        print(f"❌ Telegram測試失敗: {e}")
        return False

def test_database():
    """測試資料庫功能"""
    print("🔍 測試資料庫功能...")
    
    try:
        from database_manager import TaifexDatabaseManager
        
        db_manager = TaifexDatabaseManager()
        
        # 檢查是否有資料
        recent_data = db_manager.get_recent_data(7)
        
        if not recent_data.empty:
            print(f"✅ 資料庫包含 {len(recent_data)} 筆最近7天的資料")
            print(f"   最新資料日期: {recent_data['date'].max()}")
            print(f"   最舊資料日期: {recent_data['date'].min()}")
            return True
        else:
            print("⚠️ 資料庫中沒有最近7天的資料")
            
            # 檢查是否有任何資料
            all_data = db_manager.get_recent_data(365)
            if not all_data.empty:
                print(f"✅ 資料庫包含 {len(all_data)} 筆歷史資料")
                return True
            else:
                print("❌ 資料庫完全沒有資料")
                return False
                
    except ImportError:
        print("❌ 缺少資料庫相關模組")
        return False
    except Exception as e:
        print(f"❌ 資料庫測試失敗: {e}")
        return False

def test_chart_generation():
    """測試圖表生成功能"""
    print("🔍 測試圖表生成功能...")
    
    try:
        from chart_generator import ChartGenerator
        from database_manager import TaifexDatabaseManager
        
        # 取得測試資料
        db_manager = TaifexDatabaseManager()
        test_data = db_manager.get_recent_data(30)
        
        if test_data.empty:
            print("⚠️ 沒有資料可用於圖表生成測試")
            return False
        
        chart_generator = ChartGenerator(output_dir="test_charts")
        
        # 嘗試生成一個簡單圖表
        chart_paths = chart_generator.generate_all_charts(test_data)
        
        if chart_paths:
            print(f"✅ 成功生成 {len(chart_paths)} 個圖表")
            for path in chart_paths[:3]:  # 只顯示前3個
                print(f"   📊 {path}")
            return True
        else:
            print("❌ 圖表生成失敗")
            return False
            
    except ImportError:
        print("❌ 缺少圖表生成相關模組")
        return False
    except Exception as e:
        print(f"❌ 圖表生成測試失敗: {e}")
        return False

def main():
    """主測試函數"""
    print("=" * 60)
    print("🧪 台期所爬蟲整合功能測試")
    print("=" * 60)
    print(f"測試時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 執行各項測試
    tests = [
        ("資料庫功能", test_database),
        ("Google Sheets整合", test_google_sheets),
        ("Telegram整合", test_telegram),
        ("圖表生成功能", test_chart_generation),
    ]
    
    results = {}
    
    for test_name, test_func in tests:
        print(f"\n--- {test_name} ---")
        try:
            results[test_name] = test_func()
        except Exception as e:
            print(f"❌ {test_name}測試出現異常: {e}")
            results[test_name] = False
        print()
    
    # 總結報告
    print("=" * 60)
    print("📋 測試結果總結")
    print("=" * 60)
    
    passed = 0
    total = len(results)
    
    for test_name, result in results.items():
        status = "✅ 通過" if result else "❌ 失敗"
        print(f"{test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\n總計: {passed}/{total} 項測試通過")
    
    if passed == total:
        print("🎉 所有整合功能正常！")
        return 0
    else:
        print("⚠️ 部分功能需要設定或修正")
        print("\n💡 建議:")
        
        if not results.get("Google Sheets整合", False):
            print("- 請依照 GOOGLE_SHEETS_SETUP.md 設定Google Sheets")
        
        if not results.get("Telegram整合", False):
            print("- 請依照 TELEGRAM_SETUP.md 設定Telegram Bot")
        
        if not results.get("資料庫功能", False):
            print("- 請先執行爬蟲取得一些資料")
        
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code) 