#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試Telegram派報功能
從歷史資料中取得最近30天資料生成圖表
"""

import logging
from google_sheets_manager import GoogleSheetsManager
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import matplotlib.font_manager as fm

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_telegram_report():
    """建立Telegram派報"""
    print("📊 開始建立Telegram派報...")
    
    # 連接Google Sheets
    manager = GoogleSheetsManager()
    
    if not manager.client:
        print("❌ Google Sheets連接失敗")
        return
    
    # 連接到現有試算表
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1Ltv8zsQcCQ5SiaYKsgCDetNC-SqEMZP4V33S2nKMuWI"
    spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
    
    if not spreadsheet:
        print("❌ 連接試算表失敗")
        return
    
    print(f"✅ 成功連接試算表: {spreadsheet.title}")
    
    # 取得最近30天資料
    print("\n📊 取得最近30天資料...")
    recent_data = manager.get_recent_data_for_report(30)
    
    print(f"📊 取得資料筆數: {len(recent_data)}")
    
    if recent_data.empty:
        print("⚠️ 沒有最近30天的資料，使用所有可用資料")
        # 如果沒有最近30天資料，直接從工作表取得所有資料
        try:
            worksheet = manager.spreadsheet.worksheet("歷史資料")
            all_data = worksheet.get_all_records()
            recent_data = pd.DataFrame(all_data)
            print(f"📊 使用所有資料: {len(recent_data)} 筆")
        except Exception as e:
            print(f"❌ 取得資料失敗: {e}")
            return
    
    if recent_data.empty:
        print("❌ 沒有可用的資料")
        return
    
    # 分析資料
    print("\n🔍 分析最新資料...")
    
    # 篩選最新日期的資料（假設最新就是今天的資料）
    if '日期' in recent_data.columns:
        try:
            recent_data['日期'] = pd.to_datetime(recent_data['日期'], errors='coerce')
            latest_date = recent_data['日期'].max()
            latest_data = recent_data[recent_data['日期'] == latest_date]
            print(f"📅 最新資料日期: {latest_date.strftime('%Y-%m-%d') if pd.notna(latest_date) else '未知'}")
        except:
            # 如果日期處理失敗，使用所有資料
            latest_data = recent_data
            print("📅 使用所有可用資料")
    else:
        latest_data = recent_data
        print("📅 使用所有可用資料")
    
    # 顯示三大法人部位摘要
    print("\n📊 三大法人部位摘要:")
    summary_text = []
    
    for _, row in latest_data.iterrows():
        identity = row.get('身份別', '未知')
        long_pos = row.get('多方未平倉口數', 0)
        short_pos = row.get('空方未平倉口數', 0)
        
        try:
            long_pos = int(float(str(long_pos).replace(',', ''))) if long_pos else 0
            short_pos = int(float(str(short_pos).replace(',', ''))) if short_pos else 0
            net_pos = long_pos - short_pos
            
            print(f"  {identity}: 多方{long_pos:,}口 vs 空方{short_pos:,}口 (淨{'+' if net_pos >= 0 else ''}{net_pos:,}口)")
            summary_text.append(f"{identity}: 淨{'+' if net_pos >= 0 else ''}{net_pos:,}口")
            
        except Exception as e:
            print(f"  {identity}: 資料格式錯誤")
    
    # 生成簡單圖表
    print("\n📈 生成圖表...")
    
    try:
        # 準備圖表資料
        identities = []
        net_positions = []
        
        for _, row in latest_data.iterrows():
            identity = row.get('身份別', '未知')
            long_pos = row.get('多方未平倉口數', 0)
            short_pos = row.get('空方未平倉口數', 0)
            
            try:
                long_pos = int(float(str(long_pos).replace(',', ''))) if long_pos else 0
                short_pos = int(float(str(short_pos).replace(',', ''))) if short_pos else 0
                net_pos = long_pos - short_pos
                
                identities.append(identity)
                net_positions.append(net_pos)
                
            except:
                continue
        
        if identities and net_positions:
            # 建立圖表
            fig, ax = plt.subplots(figsize=(10, 6))
            bars = ax.bar(identities, net_positions, 
                         color=['red' if x < 0 else 'green' for x in net_positions])
            
            ax.set_title('三大法人未平倉淨部位', fontsize=16, fontweight='bold')
            ax.set_ylabel('淨部位 (口)', fontsize=12)
            ax.axhline(y=0, color='black', linestyle='-', alpha=0.3)
            
            # 在柱狀圖上顯示數值
            for bar, value in zip(bars, net_positions):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + (100 if height >= 0 else -200),
                       f'{value:,}', ha='center', va='bottom' if height >= 0 else 'top')
            
            plt.tight_layout()
            
            # 儲存圖表
            chart_filename = f"charts/telegram_report_{datetime.now().strftime('%Y%m%d')}.png"
            plt.savefig(chart_filename, dpi=300, bbox_inches='tight')
            plt.close()
            
            print(f"✅ 圖表已儲存: {chart_filename}")
            
        else:
            print("⚠️ 沒有足夠的資料生成圖表")
    
    except Exception as e:
        print(f"❌ 圖表生成失敗: {e}")
    
    # 生成文字報告
    report_text = f"""
📊 **台期所三大法人分析**
📅 日期: {datetime.now().strftime('%Y-%m-%d')}

💼 **未平倉淨部位:**
{chr(10).join(summary_text)}

📈 **投資建議:**
• 外資淨部位變化通常反映國際資金流向
• 投信動向常代表內資法人看法
• 自營商部位變化較為短期導向

🔗 詳細資料: https://docs.google.com/spreadsheets/d/1Ltv8zsQcCQ5SiaYKsgCDetNC-SqEMZP4V33S2nKMuWI
    """
    
    print("\n📝 Telegram報告內容:")
    print(report_text)
    
    # 儲存報告
    report_filename = f"reports/telegram_report_{datetime.now().strftime('%Y%m%d')}.txt"
    try:
        import os
        os.makedirs('reports', exist_ok=True)
        with open(report_filename, 'w', encoding='utf-8') as f:
            f.write(report_text)
        print(f"✅ 報告已儲存: {report_filename}")
    except Exception as e:
        print(f"⚠️ 報告儲存失敗: {e}")

if __name__ == "__main__":
    create_telegram_report() 