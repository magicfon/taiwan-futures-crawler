#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試遺漏資料檢查工作流程
模擬GitHub Actions的執行流程
"""

import subprocess
import sys
import os
from datetime import datetime

def test_missing_check_workflow():
    """測試完整的遺漏資料檢查工作流程"""
    print("🧪 開始測試遺漏資料檢查工作流程...")
    print("=" * 60)
    
    # 步驟1: 檢查並補齊遺漏資料
    print("📋 步驟1: 檢查並補齊遺漏資料")
    try:
        result = subprocess.run([
            'python', 'check_and_fill_missing_data.py',
            '--days', '30'
        ], capture_output=True, text=True, timeout=600)
        
        if result.returncode == 0:
            print("✅ 遺漏資料檢查完成")
        else:
            print("⚠️ 遺漏資料檢查有警告，但繼續執行")
            
        # 顯示輸出（最後20行）
        if result.stdout:
            lines = result.stdout.strip().split('\n')
            print("輸出摘要（最後20行）:")
            for line in lines[-20:]:
                print(f"  {line}")
                
    except subprocess.TimeoutExpired:
        print("❌ 遺漏資料檢查超時")
        return False
    except Exception as e:
        print(f"❌ 遺漏資料檢查失敗: {e}")
        return False
    
    print("\n" + "=" * 60)
    
    # 步驟2: 執行今日爬蟲
    print("📋 步驟2: 執行今日爬蟲")
    try:
        result = subprocess.run([
            'python', 'taifex_crawler.py',
            '--date-range', 'today',
            '--contracts', 'TX,TE,MTX'
        ], capture_output=True, text=True, timeout=300)
        
        if result.returncode == 0:
            print("✅ 今日爬蟲執行成功")
        else:
            print("⚠️ 今日爬蟲執行有警告")
            
    except subprocess.TimeoutExpired:
        print("❌ 今日爬蟲超時")
    except Exception as e:
        print(f"❌ 今日爬蟲失敗: {e}")
    
    print("\n" + "=" * 60)
    
    # 步驟3: 檢查生成的檔案
    print("📋 步驟3: 檢查生成的檔案")
    
    files_to_check = [
        'reports/missing_data_check_report.json',
        'missing_data_check.log',
        'taifex_crawler.log'
    ]
    
    for file_path in files_to_check:
        if os.path.exists(file_path):
            size = os.path.getsize(file_path)
            print(f"✅ {file_path} ({size:,} bytes)")
        else:
            print(f"❌ {file_path} (不存在)")
    
    # 檢查輸出目錄
    if os.path.exists('output'):
        output_files = os.listdir('output')
        print(f"📁 output/ 目錄: {len(output_files)} 個檔案")
        for file in output_files[:5]:  # 只顯示前5個
            print(f"   - {file}")
        if len(output_files) > 5:
            print(f"   ... 還有 {len(output_files)-5} 個檔案")
    else:
        print("📁 output/ 目錄: 不存在")
    
    print("\n" + "=" * 60)
    
    # 步驟4: 檢查資料庫狀態
    print("📋 步驟4: 檢查資料庫狀態")
    try:
        result = subprocess.run([
            'python', 'check_db_data.py'
        ], capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            print("✅ 資料庫狀態檢查完成")
            # 顯示最後幾行輸出
            if result.stdout:
                lines = result.stdout.strip().split('\n')
                for line in lines[-10:]:
                    if line.strip():
                        print(f"  {line}")
        else:
            print("⚠️ 資料庫狀態檢查有問題")
            
    except Exception as e:
        print(f"❌ 資料庫狀態檢查失敗: {e}")
    
    print("\n" + "=" * 60)
    
    # 步驟5: 生成執行摘要
    print("📋 步驟5: 生成執行摘要")
    
    summary = {
        'test_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'test_status': 'completed',
        'files_generated': []
    }
    
    for file_path in files_to_check:
        if os.path.exists(file_path):
            summary['files_generated'].append(file_path)
    
    # 讀取遺漏資料檢查報告
    if os.path.exists('reports/missing_data_check_report.json'):
        try:
            import json
            with open('reports/missing_data_check_report.json', 'r', encoding='utf-8') as f:
                missing_report = json.load(f)
            
            print("📊 遺漏資料檢查摘要:")
            print(f"   檢查期間: {missing_report.get('check_period_days', 'N/A')}天")
            print(f"   遺漏日期數: {missing_report.get('missing_dates_count', 'N/A')}")
            print(f"   資料庫狀態: {missing_report.get('database_status', 'N/A')}")
            print(f"   Google Sheets: {missing_report.get('google_sheets_status', 'N/A')}")
            
        except Exception as e:
            print(f"⚠️ 無法讀取遺漏資料檢查報告: {e}")
    
    print("\n🎉 測試工作流程完成！")
    return True

def main():
    """主程式"""
    print("🧪 測試遺漏資料檢查工作流程")
    print(f"📅 測試時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("🔄 模擬GitHub Actions執行流程...\n")
    
    try:
        success = test_missing_check_workflow()
        
        if success:
            print("\n✅ 所有測試步驟完成")
            sys.exit(0)
        else:
            print("\n❌ 測試過程中發生錯誤")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n⚠️ 測試被使用者中斷")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 測試失敗: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 