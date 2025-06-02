#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
GitHub Actions 自動化設定腳本
幫助快速檢查和設定GitHub Actions環境
"""

import os
import sys
import subprocess
import json
from pathlib import Path

def check_git_repository():
    """檢查是否在Git repository中"""
    try:
        result = subprocess.run(['git', 'rev-parse', '--git-dir'], 
                              capture_output=True, text=True, check=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False

def check_remote_repository():
    """檢查是否有遠端repository"""
    try:
        result = subprocess.run(['git', 'remote', '-v'], 
                              capture_output=True, text=True, check=True)
        return 'origin' in result.stdout
    except subprocess.CalledProcessError:
        return False

def check_required_files():
    """檢查必要檔案是否存在"""
    required_files = [
        'taifex_crawler.py',
        'requirements.txt',
        '.github/workflows/daily_crawler.yml'
    ]
    
    missing_files = []
    for file_path in required_files:
        if not Path(file_path).exists():
            missing_files.append(file_path)
    
    return missing_files

def create_directories():
    """建立必要的目錄"""
    directories = ['output', '.github/workflows']
    
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
        print(f"✅ 建立目錄: {directory}")

def test_python_script():
    """測試Python腳本是否能正常執行"""
    try:
        result = subprocess.run([sys.executable, 'taifex_crawler.py', '--help'], 
                              capture_output=True, text=True, timeout=10)
        return result.returncode == 0
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
        return False

def show_setup_status():
    """顯示設定狀態"""
    print("🔍 檢查GitHub Actions自動化設定狀態...")
    print("=" * 50)
    
    # 1. 檢查Git repository
    if check_git_repository():
        print("✅ Git repository: 已初始化")
    else:
        print("❌ Git repository: 未初始化")
        print("   請執行: git init")
    
    # 2. 檢查遠端repository
    if check_remote_repository():
        print("✅ 遠端repository: 已設定")
    else:
        print("❌ 遠端repository: 未設定")
        print("   請執行: git remote add origin <your-repo-url>")
    
    # 3. 檢查必要檔案
    missing_files = check_required_files()
    if not missing_files:
        print("✅ 必要檔案: 完整")
    else:
        print("❌ 缺少檔案:")
        for file in missing_files:
            print(f"   - {file}")
    
    # 4. 測試Python腳本
    if test_python_script():
        print("✅ Python腳本: 可正常執行")
    else:
        print("❌ Python腳本: 執行失敗")
        print("   請檢查相依套件是否已安裝")
    
    # 5. 檢查工作流程檔案
    workflow_file = Path('.github/workflows/daily_crawler.yml')
    if workflow_file.exists():
        print("✅ GitHub Actions工作流程: 已設定")
    else:
        print("❌ GitHub Actions工作流程: 未設定")

def show_next_steps():
    """顯示後續步驟"""
    print("\n🚀 後續步驟:")
    print("=" * 50)
    
    steps = [
        "1. 將專案推送到GitHub:",
        "   git add .",
        "   git commit -m 'Setup GitHub Actions automation'",
        "   git push -u origin main",
        "",
        "2. 在GitHub repository中啟用Actions:",
        "   - 進入repository的Settings",
        "   - 點選Actions -> General",
        "   - 確保Allow all actions and reusable workflows已勾選",
        "",
        "3. 手動測試工作流程:",
        "   - 進入Actions標籤",
        "   - 選擇'每日台期所資料爬取'",
        "   - 點選'Run workflow'進行測試",
        "",
        "4. 檢查執行結果:",
        "   - 查看Actions執行日誌",
        "   - 確認output目錄中有產生檔案",
        "   - 檢查是否有自動commit"
    ]
    
    for step in steps:
        print(step)

def main():
    """主程序"""
    print("🤖 GitHub Actions 自動化設定助手")
    print("=" * 50)
    
    # 建立必要目錄
    create_directories()
    
    # 顯示設定狀態
    show_setup_status()
    
    # 顯示後續步驟
    show_next_steps()
    
    print("\n📚 更多資訊請參考: AUTOMATION_README.md")

if __name__ == "__main__":
    main() 