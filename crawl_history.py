#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
歷史資料爬取腳本
從指定日期範圍爬取台期所資料，並上傳到Google Sheets歷史資料庫
"""

from taifex_crawler import TaifexCrawler
from google_sheets_manager import GoogleSheetsManager
import pandas as pd
from datetime import datetime, timedelta
import logging
import argparse

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def crawl_historical_data(start_date, end_date=None, contracts=None, identities=None):
    """
    爬取歷史資料並上傳到Google Sheets
    
    Args:
        start_date: 開始日期 (格式: 'YYYY-MM-DD')
        end_date: 結束日期 (格式: 'YYYY-MM-DD')，預設為今天
        contracts: 契約列表，預設為 ['TX', 'TE', 'MTX', 'ZMX', 'NQF']
        identities: 身份別列表，預設為 ['自營商', '投信', '外資']
    """
    if not end_date:
        end_date = datetime.now().strftime('%Y-%m-%d')
    
    if not contracts:
        contracts = ['TX', 'TE', 'MTX', 'ZMX', 'NQF']  # 包含所有主要期貨契約
    
    if not identities:
        identities = ['自營商', '投信', '外資']
    
    print(f"🚀 開始爬取歷史資料...")
    print(f"📅 日期範圍: {start_date} 至 {end_date}")
    print(f"📊 契約: {', '.join(contracts)}")
    print(f"👥 身份別: {', '.join(identities)}")
    
    # 初始化爬蟲
    crawler = TaifexCrawler(output_dir="output_history", delay=1.0)
    
    # 轉換日期格式為datetime物件
    try:
        start_date_obj = datetime.strptime(start_date, '%Y-%m-%d')
        end_date_obj = datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError as e:
        print(f"❌ 日期格式錯誤: {e}")
        return pd.DataFrame()
    
    # 爬取資料
    all_data = []
    
    for contract in contracts:
        print(f"\n📈 正在爬取 {contract} 契約資料...")
        
        try:
            # 爬取日期範圍的資料
            contract_data = crawler.crawl_date_range(
                start_date=start_date_obj,
                end_date=end_date_obj,
                contracts=[contract],
                identities=identities
            )
            
            if contract_data is not None and not contract_data.empty:
                print(f"✅ {contract} 爬取成功: {len(contract_data)} 筆資料")
                all_data.append(contract_data)
            else:
                print(f"⚠️ {contract} 沒有資料")
                
        except Exception as e:
            print(f"❌ {contract} 爬取失敗: {e}")
            logger.error(f"爬取 {contract} 失敗: {e}")
    
    # 合併所有資料
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        print(f"\n📊 總共爬取到 {len(combined_df)} 筆資料")
        
        # 儲存本地CSV
        filename = f"output_history/taifex_history_{start_date}_{end_date}.csv"
        combined_df.to_csv(filename, index=False, encoding='utf-8-sig')
        print(f"💾 本地備份已儲存: {filename}")
        
        # 上傳到Google Sheets
        print(f"\n🌐 正在上傳到Google Sheets...")
        
        try:
            # 連接Google Sheets
            manager = GoogleSheetsManager()
            
            if not manager.client:
                print("❌ Google Sheets連接失敗")
                return combined_df
            
            # 連接到試算表
            spreadsheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
            spreadsheet = manager.connect_spreadsheet(spreadsheet_url)
            
            if not spreadsheet:
                print("❌ 連接試算表失敗")
                return combined_df
            
            # 檢查歷史資料工作表是否存在
            worksheet_names = [ws.title for ws in spreadsheet.worksheets()]
            
            if "歷史資料" not in worksheet_names:
                print("⚠️ 歷史資料工作表不存在，正在建立...")
                try:
                    worksheet = spreadsheet.add_worksheet(title="歷史資料", rows=10000, cols=20)
                    headers = manager.get_history_headers()
                    worksheet.update('A1', [headers])
                    worksheet.format('A1:Z1', {
                        'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.9},
                        'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}
                    })
                    print("✅ 歷史資料工作表建立成功!")
                except Exception as e:
                    print(f"❌ 建立歷史資料工作表失敗: {e}")
                    return combined_df
            
            # 上傳資料
            success = manager.upload_data(combined_df, "歷史資料")
            
            if success:
                print("✅ 資料已成功上傳到Google Sheets!")
                print(f"🌐 查看結果: {manager.get_spreadsheet_url()}")
            else:
                print("❌ 上傳到Google Sheets失敗")
            
        except Exception as e:
            print(f"❌ Google Sheets處理失敗: {e}")
            logger.error(f"Google Sheets處理失敗: {e}")
        
        return combined_df
    
    else:
        print("❌ 沒有爬取到任何資料")
        return pd.DataFrame()

def main():
    """主程式入口"""
    parser = argparse.ArgumentParser(description='台期所歷史資料爬取工具')
    parser.add_argument('--start-date', '-s', type=str, default='2025-05-01',
                        help='開始日期 (格式: YYYY-MM-DD)，預設為 2025-05-01')
    parser.add_argument('--end-date', '-e', type=str, default=None,
                        help='結束日期 (格式: YYYY-MM-DD)，預設為今天')
    parser.add_argument('--contracts', '-c', nargs='+', default=['TX', 'TE', 'MTX', 'ZMX', 'NQF'],
                        help='契約代碼清單，預設為 TX,TE,MTX,ZMX,NQF')
    parser.add_argument('--identities', '-i', nargs='+', 
                        default=['自營商', '投信', '外資'],
                        help='身份別清單，預設為全部')
    
    args = parser.parse_args()
    
    # 執行爬取
    result_df = crawl_historical_data(
        start_date=args.start_date,
        end_date=args.end_date,
        contracts=args.contracts,
        identities=args.identities
    )
    
    if not result_df.empty:
        print(f"\n🎉 歷史資料爬取完成!")
        print(f"📊 總筆數: {len(result_df)}")
        print(f"📅 日期範圍: {result_df['日期'].min()} 至 {result_df['日期'].max()}")
        
        # 顯示摘要統計
        print(f"\n📋 資料摘要:")
        for identity in result_df['身份別'].unique():
            identity_data = result_df[result_df['身份別'] == identity]
            print(f"  {identity}: {len(identity_data)} 筆資料")
    else:
        print("😞 沒有爬取到資料")

if __name__ == "__main__":
    main() 