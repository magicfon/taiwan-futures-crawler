#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
測試Telegram通知功能
模擬兩階段的Telegram通知：
1. 下午2點：發送交易量摘要
2. 下午3點半：發送圖表報告
"""

import pandas as pd
import datetime
import pytz
import os
from pathlib import Path

# 確保能夠導入模組
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from telegram_notifier import TelegramNotifier
from taifex_crawler import generate_trading_summary

TW_TZ = pytz.timezone('Asia/Taipei')

def create_sample_trading_data():
    """創建示例交易量資料"""
    current_date = datetime.datetime.now(TW_TZ).strftime('%Y/%m/%d')
    
    sample_data = [
        {
            '日期': current_date,
            '契約名稱': 'TX',
            '身份別': '自營商',
            '多方交易口數': 45000,
            '多方契約金額': 85000000,
            '空方交易口數': 38000,
            '空方契約金額': 72000000,
            '多空淨額交易口數': 7000,
            '多空淨額契約金額': 13000000,
        },
        {
            '日期': current_date,
            '契約名稱': 'TX',
            '身份別': '投信',
            '多方交易口數': 12000,
            '多方契約金額': 22000000,
            '空方交易口數': 15000,
            '空方契約金額': 28000000,
            '多空淨額交易口數': -3000,
            '多空淨額契約金額': -6000000,
        },
        {
            '日期': current_date,
            '契約名稱': 'TX',
            '身份別': '外資',
            '多方交易口數': 25000,
            '多方契約金額': 47000000,
            '空方交易口數': 30000,
            '空方契約金額': 56000000,
            '多空淨額交易口數': -5000,
            '多空淨額契約金額': -9000000,
        },
        {
            '日期': current_date,
            '契約名稱': 'TE',
            '身份別': '自營商',
            '多方交易口數': 8000,
            '多方契約金額': 12000000,
            '空方交易口數': 6000,
            '空方契約金額': 9000000,
            '多空淨額交易口數': 2000,
            '多空淨額契約金額': 3000000,
        },
        {
            '日期': current_date,
            '契約名稱': 'MTX',
            '身份別': '外資',
            '多方交易口數': 15000,
            '多方契約金額': 8000000,
            '空方交易口數': 18000,
            '空方契約金額': 9500000,
            '多空淨額交易口數': -3000,
            '多空淨額契約金額': -1500000,
        }
    ]
    
    return pd.DataFrame(sample_data)

def test_trading_summary_notification():
    """測試交易量摘要通知（下午2點模式）"""
    print("=" * 60)
    print("🧪 測試交易量摘要通知（下午2點模式）")
    print("=" * 60)
    
    # 創建示例資料
    df = create_sample_trading_data()
    current_time = datetime.datetime.now(TW_TZ)
    
    # 生成摘要文字
    summary_text = generate_trading_summary(df, current_time)
    print("\n📝 生成的摘要內容：")
    print("-" * 40)
    print(summary_text)
    print("-" * 40)
    
    # 初始化Telegram通知器
    notifier = TelegramNotifier()
    
    if notifier.is_configured():
        print("\n📱 正在發送Telegram通知...")
        success = notifier.send_simple_message(summary_text)
        
        if success:
            print("✅ 交易量摘要已成功發送到Telegram")
        else:
            print("❌ Telegram發送失敗")
    else:
        print("⚠️ Telegram未配置，跳過發送測試")
        print("💡 提示：請設定環境變數 TELEGRAM_BOT_TOKEN 和 TELEGRAM_CHAT_ID")

def test_chart_report_notification():
    """測試圖表報告通知（下午3點半模式）"""
    print("\n" + "=" * 60)
    print("🧪 測試圖表報告通知（下午3點半模式）")
    print("=" * 60)
    
    # 檢查是否有圖表檔案
    charts_dir = Path("charts")
    chart_files = []
    
    if charts_dir.exists():
        chart_files = list(charts_dir.glob("*.png"))
    
    if not chart_files:
        print("⚠️ 沒有找到圖表檔案")
        print("💡 請先執行完整資料爬取以生成圖表")
        return
    
    print(f"📊 找到 {len(chart_files)} 個圖表檔案：")
    for chart in chart_files:
        print(f"  • {chart.name}")
    
    # 生成圖表報告摘要
    summary_text = f"""🎯 台期所完整資料分析報告
⏰ 報告時間: {datetime.datetime.now(TW_TZ).strftime('%Y/%m/%d %H:%M')}

📊 本次分析包含：
• 三大法人持倉分析
• 多空趨勢圖表
• 成交量與未平倉變化

🔄 資料來源：台灣期貨交易所
📈 分析期間：最近30個交易日

💡 以下為詳細圖表分析："""
    
    print("\n📝 圖表報告摘要：")
    print("-" * 40)
    print(summary_text)
    print("-" * 40)
    
    # 初始化Telegram通知器
    notifier = TelegramNotifier()
    
    if notifier.is_configured():
        print("\n📱 正在發送圖表報告到Telegram...")
        chart_paths = [str(chart) for chart in chart_files[:3]]  # 限制前3個圖表避免發送過多
        success = notifier.send_chart_report(chart_paths, summary_text)
        
        if success:
            print("✅ 圖表報告已成功發送到Telegram")
        else:
            print("❌ 圖表報告發送部分失敗")
    else:
        print("⚠️ Telegram未配置，跳過發送測試")
        print("💡 提示：請設定環境變數 TELEGRAM_BOT_TOKEN 和 TELEGRAM_CHAT_ID")

def main():
    """主測試函數"""
    print("🚀 台期所Telegram通知功能測試")
    print(f"⏰ 測試時間：{datetime.datetime.now(TW_TZ).strftime('%Y/%m/%d %H:%M:%S')}")
    
    # 檢查環境變數
    print(f"\n🔧 環境變數檢查：")
    print(f"TELEGRAM_BOT_TOKEN: {'✅ 已設定' if os.environ.get('TELEGRAM_BOT_TOKEN') else '❌ 未設定'}")
    print(f"TELEGRAM_CHAT_ID: {'✅ 已設定' if os.environ.get('TELEGRAM_CHAT_ID') else '❌ 未設定'}")
    
    # 測試交易量摘要通知
    test_trading_summary_notification()
    
    # 測試圖表報告通知
    test_chart_report_notification()
    
    print("\n" + "=" * 60)
    print("🎉 測試完成！")
    print("💡 GitHub Actions工作流程將會在以下時間自動執行：")
    print("   • 下午2:00 - 交易量資料爬取 + 文字摘要通知")
    print("   • 下午3:30 - 完整資料爬取 + 圖表報告通知")
    print("=" * 60)

if __name__ == "__main__":
    main() 