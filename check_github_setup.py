#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
GitHub Actions 設定檢查工具
檢查是否已正確設定所有必要的配置
"""

import os
import json
import sys
from pathlib import Path

def check_workflow_files():
    """檢查workflow檔案是否存在"""
    print("🔍 檢查 GitHub Actions Workflow 檔案...")
    
    workflow_dir = Path(".github/workflows")
    required_workflows = [
        "crawl_trading_data.yml",
        "crawl_complete_data.yml", 
        "manual_crawl.yml"
    ]
    
    missing_files = []
    
    for workflow in required_workflows:
        workflow_path = workflow_dir / workflow
        if workflow_path.exists():
            print(f"  ✅ {workflow}")
        else:
            print(f"  ❌ {workflow} - 檔案不存在")
            missing_files.append(workflow)
    
    return len(missing_files) == 0

def check_secrets_template():
    """檢查secrets設定模板"""
    print("\n🔑 檢查環境變數設定...")
    
    # 檢查是否在GitHub Actions環境中
    if os.environ.get('GITHUB_ACTIONS') == 'true':
        print("  🤖 運行在 GitHub Actions 環境中")
        
        # 檢查必要的secrets
        required_secrets = ['GOOGLE_SHEETS_CREDENTIALS']
        optional_secrets = ['TELEGRAM_BOT_TOKEN', 'TELEGRAM_CHAT_ID']
        
        for secret in required_secrets:
            if os.environ.get(secret):
                print(f"  ✅ {secret} - 已設定")
            else:
                print(f"  ❌ {secret} - 未設定 (必要)")
        
        for secret in optional_secrets:
            if os.environ.get(secret):
                print(f"  ✅ {secret} - 已設定")
            else:
                print(f"  ⚠️ {secret} - 未設定 (可選)")
                
        return bool(os.environ.get('GOOGLE_SHEETS_CREDENTIALS'))
    else:
        print("  ℹ️ 不在 GitHub Actions 環境中")
        print("  💡 請確保在 GitHub Repository 的 Settings → Secrets 中設定:")
        print("     - GOOGLE_SHEETS_CREDENTIALS (必要)")
        print("     - TELEGRAM_BOT_TOKEN (可選)")
        print("     - TELEGRAM_CHAT_ID (可選)")
        return True

def check_google_credentials():
    """檢查Google憑證格式"""
    print("\n📊 檢查 Google Sheets 憑證...")
    
    credentials_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
    
    if not credentials_json:
        print("  ⚠️ GOOGLE_SHEETS_CREDENTIALS 環境變數未設定")
        print("  💡 這在本地環境是正常的，請確保在GitHub Secrets中已設定")
        return True
    
    try:
        credentials = json.loads(credentials_json)
        
        required_fields = [
            'type', 'project_id', 'private_key_id', 'private_key',
            'client_email', 'client_id', 'auth_uri', 'token_uri'
        ]
        
        for field in required_fields:
            if field in credentials:
                print(f"  ✅ {field}")
            else:
                print(f"  ❌ {field} - 缺少欄位")
                return False
        
        # 檢查憑證類型
        if credentials.get('type') != 'service_account':
            print("  ❌ 憑證類型不正確，應為 'service_account'")
            return False
        
        print("  ✅ Google 服務帳號憑證格式正確")
        return True
        
    except json.JSONDecodeError:
        print("  ❌ GOOGLE_SHEETS_CREDENTIALS 不是有效的 JSON 格式")
        return False
    except Exception as e:
        print(f"  ❌ 憑證檢查失敗: {e}")
        return False

def check_requirements():
    """檢查requirements.txt"""
    print("\n📦 檢查 Python 相依套件...")
    
    requirements_file = Path("requirements.txt")
    
    if not requirements_file.exists():
        print("  ❌ requirements.txt 檔案不存在")
        return False
    
    # 檢查關鍵套件
    required_packages = [
        'requests', 'pandas', 'beautifulsoup4', 'gspread', 
        'google-auth', 'schedule', 'tqdm'
    ]
    
    try:
        with open(requirements_file, 'r', encoding='utf-8') as f:
            content = f.read().lower()
        
        for package in required_packages:
            if package.lower() in content:
                print(f"  ✅ {package}")
            else:
                print(f"  ❌ {package} - 套件未列在 requirements.txt 中")
                return False
        
        return True
        
    except Exception as e:
        print(f"  ❌ 讀取 requirements.txt 失敗: {e}")
        return False

def check_main_script():
    """檢查主要爬蟲腳本"""
    print("\n🕷️ 檢查主要爬蟲腳本...")
    
    script_file = Path("taifex_crawler.py")
    
    if not script_file.exists():
        print("  ❌ taifex_crawler.py 檔案不存在")
        return False
    
    try:
        with open(script_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 檢查關鍵功能
        checks = [
            ('data_type', 'TRADING'),
            ('data_type', 'COMPLETE'),
            ('Google Sheets', 'GoogleSheetsManager'),
            ('分階段爬取', 'DATA_TYPES')
        ]
        
        for check_name, keyword in checks:
            if keyword in content:
                print(f"  ✅ {check_name} 功能")
            else:
                print(f"  ❌ {check_name} 功能 - 找不到關鍵字: {keyword}")
        
        return True
        
    except Exception as e:
        print(f"  ❌ 檢查主要腳本失敗: {e}")
        return False

def main():
    """主要檢查程序"""
    print("🤖 GitHub Actions 台期所爬蟲設定檢查")
    print("=" * 50)
    
    checks = [
        ("Workflow檔案", check_workflow_files),
        ("環境變數", check_secrets_template), 
        ("Google憑證", check_google_credentials),
        ("Python套件", check_requirements),
        ("主要腳本", check_main_script)
    ]
    
    all_passed = True
    results = []
    
    for check_name, check_func in checks:
        try:
            result = check_func()
            results.append((check_name, result))
            if not result:
                all_passed = False
        except Exception as e:
            print(f"  ❌ 檢查 {check_name} 時發生錯誤: {e}")
            results.append((check_name, False))
            all_passed = False
    
    # 輸出總結
    print("\n" + "=" * 50)
    print("📋 檢查結果總結:")
    
    for check_name, result in results:
        status = "✅ 通過" if result else "❌ 失敗"
        print(f"  {check_name}: {status}")
    
    if all_passed:
        print("\n🎉 所有檢查都通過！系統已準備就緒。")
        print("\n📝 接下來的步驟:")
        print("  1. 確認 GitHub Secrets 已正確設定")
        print("  2. 推送程式碼到 GitHub Repository")
        print("  3. 在 Actions 頁籤中手動測試 workflow")
        print("  4. 等待自動排程執行")
        return 0
    else:
        print("\n⚠️ 部分檢查未通過，請根據上述提示進行修正。")
        print("\n🆘 如需協助，請參考 GitHub_Actions_使用說明.md")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 