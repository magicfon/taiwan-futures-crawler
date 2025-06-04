#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
處理身分別資料的圖表生成器
支援加總三種身分別或分別顯示
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import logging
import os
from datetime import datetime, timedelta
import seaborn as sns

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("身分別圖表生成")

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def load_and_process_identity_data(aggregate=True):
    """
    載入並處理身分別資料
    
    Args:
        aggregate (bool): True=加總三種身分別，False=分別處理
    """
    try:
        from google_sheets_manager import GoogleSheetsManager
        
        logger.info("🔍 載入Google Sheets身分別資料...")
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        ws = spreadsheet.worksheet('歷史資料')
        
        # 獲取所有資料
        all_values = ws.get_all_values()
        headers = all_values[0]
        data_rows = all_values[1:]
        
        df = pd.DataFrame(data_rows, columns=headers)
        
        # 過濾空行
        df = df[df['日期'].str.strip() != '']
        
        # 轉換日期
        df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d')
        
        # 轉換數值欄位
        numeric_cols = ['多空淨額交易口數', '多空淨額未平倉口數']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 過濾最近30個工作日
        end_date = datetime.now()
        start_date = end_date - timedelta(days=45)
        
        df_filtered = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
        df_filtered = df_filtered[df_filtered['日期'].dt.weekday < 5]
        
        # 取最近30個工作日
        unique_dates = sorted(df_filtered['日期'].dt.date.unique())
        if len(unique_dates) > 30:
            recent_30_dates = unique_dates[-30:]
            df_filtered = df_filtered[df_filtered['日期'].dt.date.isin(recent_30_dates)]
        
        logger.info(f"✅ 載入 {len(df_filtered)} 筆30天內的身分別資料")
        logger.info(f"📅 日期範圍: {df_filtered['日期'].min().strftime('%Y-%m-%d')} 到 {df_filtered['日期'].max().strftime('%Y-%m-%d')}")
        logger.info(f"🏷️ 身份別: {df_filtered['身份別'].unique().tolist()}")
        logger.info(f"📈 契約: {df_filtered['契約名稱'].unique().tolist()}")
        
        if aggregate:
            # 加總模式：將三種身分別加總
            logger.info("📊 使用加總模式 - 將三種身分別加總為市場整體")
            
            result_df = df_filtered.groupby(['日期', '契約名稱']).agg({
                '多空淨額交易口數': 'sum',
                '多空淨額未平倉口數': 'sum'
            }).reset_index()
            
            result_df['身份別'] = '整體市場'
            
        else:
            # 分別模式：保持三種身分別分開
            logger.info("📊 使用分別模式 - 保持三種身分別分開顯示")
            result_df = df_filtered.copy()
        
        logger.info(f"✅ 處理完成，最終有 {len(result_df)} 筆資料")
        
        return result_df
        
    except Exception as e:
        logger.error(f"❌ 資料載入失敗: {e}")
        return pd.DataFrame()

def generate_identity_charts(df, output_dir="identity_charts", separate_identities=False):
    """
    生成身分別圖表
    
    Args:
        df: 資料DataFrame
        output_dir: 輸出目錄
        separate_identities: 是否分別顯示身分別
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        chart_paths = []
        
        contracts = df['契約名稱'].unique()
        logger.info(f"🎨 開始生成 {len(contracts)} 個契約的圖表...")
        
        for contract in contracts:
            contract_data = df[df['契約名稱'] == contract].copy()
            
            if contract_data.empty:
                continue
            
            # 創建圖表
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
            
            if separate_identities:
                # 分別顯示三種身分別
                identities = contract_data['身份別'].unique()
                colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
                
                for i, identity in enumerate(identities):
                    identity_data = contract_data[contract_data['身份別'] == identity]
                    identity_data = identity_data.sort_values('日期')
                    
                    color = colors[i % len(colors)]
                    
                    # 交易量圖表
                    ax1.plot(identity_data['日期'], identity_data['多空淨額交易口數'], 
                            marker='o', linewidth=2, markersize=4, 
                            label=f'{identity}交易量', color=color, alpha=0.8)
                    
                    # 未平倉圖表
                    ax2.plot(identity_data['日期'], identity_data['多空淨額未平倉口數'], 
                            marker='s', linewidth=2, markersize=4, 
                            label=f'{identity}未平倉', color=color, alpha=0.8)
                
                title_prefix = f"{contract} - 三大法人分別"
                
            else:
                # 整體市場圖表
                contract_data = contract_data.sort_values('日期')
                
                # 交易量圖表
                ax1.plot(contract_data['日期'], contract_data['多空淨額交易口數'], 
                        marker='o', linewidth=3, markersize=5, 
                        color='#2E86AB', label='淨額交易口數')
                ax1.fill_between(contract_data['日期'], contract_data['多空淨額交易口數'], 
                               alpha=0.3, color='#2E86AB')
                
                # 未平倉圖表
                ax2.plot(contract_data['日期'], contract_data['多空淨額未平倉口數'], 
                        marker='s', linewidth=3, markersize=5, 
                        color='#A23B72', label='淨額未平倉口數')
                ax2.fill_between(contract_data['日期'], contract_data['多空淨額未平倉口數'], 
                               alpha=0.3, color='#A23B72')
                
                title_prefix = f"{contract} - 整體市場"
            
            # 設定交易量圖表
            ax1.set_title(f'{title_prefix} - 多空淨額交易口數 (30天)', fontsize=14, fontweight='bold')
            ax1.set_ylabel('交易口數', fontsize=12)
            ax1.grid(True, alpha=0.3)
            ax1.legend()
            ax1.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
            
            # 設定未平倉圖表
            ax2.set_title(f'{title_prefix} - 多空淨額未平倉口數 (30天)', fontsize=14, fontweight='bold')
            ax2.set_xlabel('日期', fontsize=12)
            ax2.set_ylabel('未平倉口數', fontsize=12)
            ax2.grid(True, alpha=0.3)
            ax2.legend()
            ax2.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
            
            # 設定日期格式
            for ax in [ax1, ax2]:
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
                ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
                plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)
            
            plt.tight_layout()
            
            # 儲存圖表
            filename = f"{contract}_{'identity_separate' if separate_identities else 'market_total'}_30days.png"
            filepath = os.path.join(output_dir, filename)
            plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            
            chart_paths.append(filepath)
            logger.info(f"✅ 已生成: {filename}")
        
        return chart_paths
        
    except Exception as e:
        logger.error(f"❌ 圖表生成失敗: {e}")
        return []

def generate_summary_text(df, is_aggregated=True):
    """生成摘要文字"""
    try:
        if df.empty:
            return "⚠️ 無資料可分析"
        
        date_range = f"{df['日期'].min().strftime('%Y-%m-%d')} 到 {df['日期'].max().strftime('%Y-%m-%d')}"
        unique_dates = df['日期'].dt.date.nunique()
        contracts = df['契約名稱'].unique().tolist()
        
        summary = f"""📊 台期所三大法人持倉分析報告 (30天)
📅 分析期間: {date_range} ({unique_dates}個交易日)
📈 分析契約: {', '.join(contracts)}

"""
        
        if is_aggregated:
            summary += "📋 本報告使用「整體市場」模式，將外資、投信、自營商三大法人數據加總\n\n"
            
            for contract in contracts:
                contract_data = df[df['契約名稱'] == contract]
                if contract_data.empty:
                    continue
                
                latest_trade = contract_data['多空淨額交易口數'].iloc[-1]
                latest_position = contract_data['多空淨額未平倉口數'].iloc[-1]
                
                trade_trend = "📈 偏多" if latest_trade > 0 else "📉 偏空" if latest_trade < 0 else "➡️ 平衡"
                position_trend = "📈 偏多" if latest_position > 0 else "📉 偏空" if latest_position < 0 else "➡️ 平衡"
                
                summary += f"""💼 {contract}契約:
   • 最新淨交易: {latest_trade:,} 口 {trade_trend}
   • 最新淨持倉: {latest_position:,} 口 {position_trend}

"""
        else:
            summary += "📋 本報告使用「分別顯示」模式，分別呈現外資、投信、自營商的數據\n\n"
            
            for contract in contracts:
                contract_data = df[df['契約名稱'] == contract]
                if contract_data.empty:
                    continue
                
                summary += f"💼 {contract}契約 (最新數據):\n"
                
                for identity in ['外資', '投信', '自營商']:
                    identity_data = contract_data[contract_data['身份別'] == identity]
                    if identity_data.empty:
                        continue
                    
                    latest_data = identity_data.iloc[-1]
                    trade_vol = latest_data['多空淨額交易口數']
                    position_vol = latest_data['多空淨額未平倉口數']
                    
                    trade_trend = "📈" if trade_vol > 0 else "📉" if trade_vol < 0 else "➡️"
                    position_trend = "📈" if position_vol > 0 else "📉" if position_vol < 0 else "➡️"
                    
                    summary += f"   • {identity}: 交易{trade_vol:,}口{trade_trend} 持倉{position_vol:,}口{position_trend}\n"
                
                summary += "\n"
        
        summary += "📱 圖表包含30天完整趨勢分析，助您掌握市場動向！"
        
        return summary
        
    except Exception as e:
        logger.error(f"❌ 摘要生成失敗: {e}")
        return "⚠️ 摘要生成失敗"

def main():
    """主程式 - 提供兩種模式選擇"""
    try:
        from telegram_notifier import TelegramNotifier
        
        print("=== 台期所三大法人持倉分析 ===\n")
        
        # 讓用戶選擇模式
        print("請選擇分析模式:")
        print("1. 整體市場模式 (三大法人加總)")
        print("2. 分別顯示模式 (三大法人分開)")
        
        choice = input("請輸入選擇 (1或2，預設為1): ").strip()
        
        if choice == "2":
            aggregate_mode = False
            separate_identities = True
            logger.info("📊 選擇模式: 分別顯示三大法人")
        else:
            aggregate_mode = True
            separate_identities = False
            logger.info("📊 選擇模式: 整體市場加總")
        
        # 1. 載入資料
        df = load_and_process_identity_data(aggregate=aggregate_mode)
        
        if df.empty:
            logger.error("❌ 無法載入資料")
            return 1
        
        # 2. 生成圖表
        chart_paths = generate_identity_charts(df, separate_identities=separate_identities)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return 1
        
        # 3. 生成摘要
        summary_text = generate_summary_text(df, is_aggregated=aggregate_mode)
        print("\n" + summary_text)
        
        # 4. 發送到Telegram
        logger.info("📱 發送到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        if notifier.test_connection():
            success = notifier.send_chart_report(chart_paths, summary_text)
            
            if success:
                logger.info("🎉 三大法人持倉分析圖表已成功發送到Telegram！")
                return 0
            else:
                logger.error("❌ Telegram發送失敗")
                return 1
        else:
            logger.error("❌ Telegram連線失敗")
            return 1
            
    except Exception as e:
        logger.error(f"❌ 執行失敗: {e}")
        return 1

if __name__ == "__main__":
    exit_code = main()
    exit(exit_code) 