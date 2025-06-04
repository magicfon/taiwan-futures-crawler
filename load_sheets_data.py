#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
從Google Sheets載入完整歷史資料並生成30天圖表
"""

import pandas as pd
import logging
import os
from datetime import datetime, timedelta

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("Google Sheets資料載入")

def load_all_sheets_data():
    """從Google Sheets載入所有歷史資料"""
    try:
        from google_sheets_manager import GoogleSheetsManager
        
        logger.info("🔍 初始化Google Sheets管理器...")
        sheets_manager = GoogleSheetsManager()
        
        if not sheets_manager.client:
            logger.error("❌ Google Sheets連線失敗")
            return pd.DataFrame()
        
        logger.info("✅ Google Sheets連線成功")
        
        # 尋找現有的試算表
        logger.info("🔍 搜尋現有的試算表...")
        
        # 嘗試從配置檔案載入試算表ID
        try:
            import json
            from pathlib import Path
            
            config_file = Path("config/spreadsheet_config.json")
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    spreadsheet_id = config.get('spreadsheet_id')
                    
                    if spreadsheet_id:
                        sheets_manager.connect_spreadsheet(spreadsheet_id)
                        logger.info(f"✅ 已連接到現有試算表: {config.get('spreadsheet_url', '')}")
        except Exception as e:
            logger.warning(f"無法載入配置檔案: {e}")
        
        # 如果沒有連接到試算表，嘗試尋找現有的
        if not sheets_manager.spreadsheet:
            logger.info("🔍 嘗試尋找包含台期所資料的試算表...")
            
            # 搜尋可能的試算表名稱
            possible_names = ["台期所資料分析", "台期所", "期貨資料", "taifex"]
            
            for name in possible_names:
                try:
                    spreadsheet = sheets_manager.client.open(name)
                    sheets_manager.spreadsheet = spreadsheet
                    logger.info(f"✅ 找到試算表: {name}")
                    break
                except:
                    continue
        
        if not sheets_manager.spreadsheet:
            logger.error("❌ 無法找到包含台期所資料的試算表")
            return pd.DataFrame()
        
        # 列出所有工作表
        worksheets = sheets_manager.spreadsheet.worksheets()
        logger.info(f"📊 試算表包含 {len(worksheets)} 個工作表:")
        for ws in worksheets:
            try:
                row_count = len(ws.get_all_values())
                logger.info(f"  • {ws.title}: {row_count} 行")
            except:
                logger.info(f"  • {ws.title}: 無法讀取")
        
        # 嘗試從不同工作表載入資料
        all_data = []
        
        # 優先順序：原始資料 > 資料 > 第一個有資料的工作表
        worksheet_priority = [
            "原始資料", "資料", "data", "台期所資料", "期貨資料",
            "Sheet1", "工作表1"
        ]
        
        found_data = False
        
        for ws_name in worksheet_priority:
            try:
                worksheet = sheets_manager.spreadsheet.worksheet(ws_name)
                logger.info(f"📖 嘗試從工作表「{ws_name}」載入資料...")
                
                # 獲取所有資料
                data = worksheet.get_all_records()
                
                if data:
                    df = pd.DataFrame(data)
                    logger.info(f"✅ 從「{ws_name}」載入了 {len(df)} 筆資料")
                    logger.info(f"📋 欄位: {list(df.columns)}")
                    
                    all_data.append(df)
                    found_data = True
                    break
                    
            except Exception as e:
                logger.debug(f"工作表「{ws_name}」不存在或無法讀取: {e}")
                continue
        
        # 如果優先工作表都沒資料，嘗試所有工作表
        if not found_data:
            for ws in worksheets:
                try:
                    logger.info(f"📖 嘗試從工作表「{ws.title}」載入資料...")
                    data = ws.get_all_records()
                    
                    if data and len(data) > 10:  # 至少要有10筆資料才算有效
                        df = pd.DataFrame(data)
                        logger.info(f"✅ 從「{ws.title}」載入了 {len(df)} 筆資料")
                        all_data.append(df)
                        found_data = True
                        break
                        
                except Exception as e:
                    logger.debug(f"工作表「{ws.title}」讀取失敗: {e}")
                    continue
        
        if not all_data:
            logger.error("❌ 所有工作表都沒有找到有效資料")
            return pd.DataFrame()
        
        # 合併所有資料
        final_df = pd.concat(all_data, ignore_index=True)
        final_df = final_df.drop_duplicates()
        
        logger.info(f"📊 總共載入 {len(final_df)} 筆資料")
        
        # 資料預處理
        if '日期' in final_df.columns:
            final_df['日期'] = pd.to_datetime(final_df['日期'], errors='coerce')
            final_df = final_df.dropna(subset=['日期'])
            
            logger.info(f"📅 日期範圍: {final_df['日期'].min()} 到 {final_df['日期'].max()}")
            
            # 計算有多少天的資料
            unique_dates = final_df['日期'].dt.date.nunique()
            logger.info(f"📆 共有 {unique_dates} 個不同日期的資料")
        
        if '契約名稱' in final_df.columns:
            contracts = final_df['契約名稱'].unique()
            logger.info(f"📈 契約類型: {list(contracts)}")
        
        return final_df
        
    except Exception as e:
        logger.error(f"❌ 載入Google Sheets資料時發生錯誤: {e}")
        return pd.DataFrame()

def generate_30day_charts_from_sheets():
    """從Google Sheets載入資料並生成30天圖表"""
    try:
        from chart_generator import ChartGenerator
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== 從Google Sheets生成30天圖表 ===\n")
        
        # 1. 載入所有歷史資料
        df = load_all_sheets_data()
        
        if df.empty:
            logger.error("❌ 無法從Google Sheets載入資料")
            return False
        
        # 2. 過濾最近30天的資料
        if '日期' in df.columns:
            end_date = datetime.now()
            start_date = end_date - timedelta(days=45)  # 稍微放寬範圍，確保有30個工作日
            
            # 過濾日期範圍
            df_filtered = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
            
            # 只保留工作日
            df_filtered = df_filtered[df_filtered['日期'].dt.weekday < 5]
            
            # 取最近30個工作日
            unique_dates = sorted(df_filtered['日期'].dt.date.unique())
            if len(unique_dates) > 30:
                recent_30_dates = unique_dates[-30:]
                df_filtered = df_filtered[df_filtered['日期'].dt.date.isin(recent_30_dates)]
            
            logger.info(f"📊 過濾後有 {len(df_filtered)} 筆30天內的資料")
            logger.info(f"📅 使用日期範圍: {df_filtered['日期'].min().strftime('%Y-%m-%d')} 到 {df_filtered['日期'].max().strftime('%Y-%m-%d')}")
            
            df = df_filtered
        
        # 3. 資料格式檢查和轉換
        logger.info("🔧 檢查資料格式...")
        
        # 確保有必要的數值欄位
        numeric_columns = ['多空淨額交易口數', '多空淨額未平倉口數']
        for col in numeric_columns:
            if col not in df.columns:
                # 嘗試從其他欄位推導
                if '多空淨額交易口數' not in df.columns and '多方交易口數' in df.columns and '空方交易口數' in df.columns:
                    df['多空淨額交易口數'] = pd.to_numeric(df['多方交易口數'], errors='coerce').fillna(0) - pd.to_numeric(df['空方交易口數'], errors='coerce').fillna(0)
                else:
                    df[col] = 0
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        logger.info(f"✅ 資料預處理完成，最終有 {len(df)} 筆有效資料")
        
        # 4. 生成圖表
        logger.info("🎨 開始生成30天圖表...")
        chart_generator = ChartGenerator(output_dir="sheets_30day_charts")
        
        chart_paths = chart_generator.generate_all_charts(df)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return False
        
        logger.info(f"✅ 已生成 {len(chart_paths)} 個圖表:")
        for path in chart_paths:
            logger.info(f"  📈 {os.path.basename(path)}")
        
        # 5. 生成摘要文字
        summary_text = chart_generator.generate_summary_text(df)
        logger.info("📝 摘要文字:")
        print(summary_text)
        
        # 6. 發送到Telegram
        logger.info("📱 發送完整30天圖表到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        # 測試連線
        if not notifier.test_connection():
            logger.error("❌ Telegram連線失敗")
            return False
        
        # 發送圖表報告
        success = notifier.send_chart_report(chart_paths, summary_text)
        
        if success:
            logger.info("🎉 完整的30天歷史圖表已成功發送到Telegram！")
            logger.info("📱 現在您應該可以看到豐富的歷史趨勢分析了")
            return True
        else:
            logger.error("❌ Telegram發送失敗")
            return False
            
    except ImportError as e:
        logger.error(f"❌ 模組載入失敗: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ 執行過程中發生錯誤: {e}")
        return False

def main():
    """主程式"""
    success = generate_30day_charts_from_sheets()
    
    if success:
        logger.info("\n🎊 任務完成！")
        logger.info("📊 完整的30天台期所持倉分析圖表已從Google Sheets載入並發送")
        logger.info("💡 現在您的Telegram應該顯示豐富的歷史趨勢資料")
        return 0
    else:
        logger.error("\n❌ 任務失敗")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 