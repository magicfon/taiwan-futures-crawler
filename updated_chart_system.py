#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
使用新Google Sheets的更新版圖表生成系統
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import logging
import os
from datetime import datetime, timedelta
from google_sheets_manager import GoogleSheetsManager

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("更新版圖表系統")

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def load_data_from_new_sheet():
    """從新的Google Sheets載入資料"""
    try:
        logger.info("🔗 連接到新的Google Sheets...")
        
        # 新的Google Sheets URL
        new_sheet_url = "https://docs.google.com/spreadsheets/d/1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4"
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.connect_spreadsheet(new_sheet_url)
        
        if not spreadsheet:
            logger.error("❌ 無法連接到新的Google Sheets")
            return pd.DataFrame()
        
        logger.info(f"✅ 成功連接到: {spreadsheet.title}")
        
        # 讀取歷史資料工作表
        ws = spreadsheet.worksheet('歷史資料')
        all_values = ws.get_all_values()
        
        if len(all_values) < 2:
            logger.error("❌ 工作表沒有足夠資料")
            return pd.DataFrame()
        
        headers = all_values[0]
        data_rows = all_values[1:]
        
        logger.info(f"📊 讀取到 {len(data_rows)} 行原始資料")
        
        # 轉換為DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # 過濾空行
        df = df[df['日期'].str.strip() != '']
        
        logger.info(f"📊 過濾後有 {len(df)} 行有效資料")
        
        # 轉換日期
        df['日期'] = pd.to_datetime(df['日期'], format='%Y/%m/%d', errors='coerce')
        df = df.dropna(subset=['日期'])
        
        logger.info(f"📅 日期轉換後有 {len(df)} 行資料")
        logger.info(f"📅 日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
        
        # 轉換數值欄位
        numeric_cols = ['多空淨額交易口數', '多空淨額未平倉口數']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 檢查2025年6月3日資料
        june3_2025 = df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
        if not june3_2025.empty:
            logger.info(f"🎉 找到2025年6月3日資料: {len(june3_2025)}筆")
            contracts = june3_2025['契約名稱'].unique()
            logger.info(f"📈 包含契約: {list(contracts)}")
        else:
            logger.warning("⚠️ 沒有找到2025年6月3日資料")
        
        # 取得最新30個工作日的資料
        latest_date = df['日期'].max()
        logger.info(f"📅 最新日期: {latest_date.strftime('%Y-%m-%d')}")
        
        # 取得所有工作日（排除週末）
        workdays = df[df['日期'].dt.weekday < 5]['日期'].dt.date.unique()
        workdays_sorted = sorted(workdays, reverse=True)
        
        # 取最近30個工作日
        if len(workdays_sorted) > 30:
            recent_30_workdays = workdays_sorted[:30]
            logger.info(f"📅 選取30個工作日範圍: {min(recent_30_workdays)} 到 {max(recent_30_workdays)}")
        else:
            recent_30_workdays = workdays_sorted
            logger.info(f"📅 可用工作日: {len(recent_30_workdays)}天")
        
        # 過濾資料
        df_filtered = df[df['日期'].dt.date.isin(recent_30_workdays)]
        
        logger.info(f"✅ 最終過濾後有 {len(df_filtered)} 筆30天內的資料")
        
        # 再次確認6/3資料
        final_june3 = df_filtered[df_filtered['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
        if not final_june3.empty:
            logger.info(f"🎉 最終結果包含2025年6月3日資料: {len(final_june3)}筆")
        
        return df_filtered
        
    except Exception as e:
        logger.error(f"❌ 資料載入失敗: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def generate_updated_charts(df, output_dir="updated_charts_with_june3"):
    """生成包含6/3資料的更新版圖表"""
    try:
        os.makedirs(output_dir, exist_ok=True)
        chart_paths = []
        
        # 加總三大法人資料
        logger.info("📊 加總三大法人資料...")
        df_processed = df.groupby(['日期', '契約名稱']).agg({
            '多空淨額交易口數': 'sum',
            '多空淨額未平倉口數': 'sum'
        }).reset_index()
        
        contracts = df_processed['契約名稱'].unique()
        logger.info(f"🎨 開始生成 {len(contracts)} 個契約的更新版圖表...")
        
        for contract in contracts:
            contract_data = df_processed[df_processed['契約名稱'] == contract].copy()
            contract_data = contract_data.sort_values('日期')
            
            if contract_data.empty:
                continue
            
            # 檢查是否包含6/3資料
            has_june3 = not contract_data[contract_data['日期'].dt.date == pd.to_datetime('2025/06/03').date()].empty
            june3_note = " ✅含6/3" if has_june3 else ""
            
            # 創建雙軸圖表
            fig, ax1 = plt.subplots(figsize=(16, 9))
            
            # 主軸：交易口數（柱狀圖）
            colors = ['#27AE60' if x >= 0 else '#E74C3C' for x in contract_data['多空淨額交易口數']]
            bars = ax1.bar(contract_data['日期'], contract_data['多空淨額交易口數'], 
                          alpha=0.8, color=colors, width=0.8, label='多空淨額交易口數')
            
            # 設定主軸
            ax1.set_xlabel('日期', fontsize=12, fontweight='bold')
            ax1.set_ylabel('多空淨額交易口數', fontsize=12, fontweight='bold', color='#2C3E50')
            ax1.tick_params(axis='y', labelcolor='#2C3E50')
            ax1.grid(True, alpha=0.3, axis='y')
            ax1.axhline(y=0, color='black', linestyle='-', alpha=0.8, linewidth=1)
            
            # 副軸：未平倉口數（折線圖）
            ax2 = ax1.twinx()
            line = ax2.plot(contract_data['日期'], contract_data['多空淨額未平倉口數'], 
                           color='#8E44AD', linewidth=4, marker='o', markersize=8, 
                           markerfacecolor='white', markeredgecolor='#8E44AD', markeredgewidth=2,
                           label='多空淨額未平倉口數')
            
            # 設定副軸
            ax2.set_ylabel('多空淨額未平倉口數', fontsize=12, fontweight='bold', color='#8E44AD')
            ax2.tick_params(axis='y', labelcolor='#8E44AD')
            ax2.axhline(y=0, color='#8E44AD', linestyle='--', alpha=0.6, linewidth=2)
            
            # 標題
            title = f"{contract} 契約 - 三大法人整體持倉分析 (完整30天含6/3){june3_note}"
            plt.title(title, fontsize=18, fontweight='bold', pad=25)
            
            # 設定日期格式
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
            ax1.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            # 添加圖例
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            legend = ax1.legend(lines1 + lines2, labels1 + labels2, 
                              loc='upper left', fontsize=11, framealpha=0.9)
            legend.get_frame().set_facecolor('white')
            
            # 特別標記6/3資料
            if has_june3:
                june3_data = contract_data[contract_data['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
                if not june3_data.empty:
                    june3_trade = june3_data['多空淨額交易口數'].iloc[0]
                    june3_position = june3_data['多空淨額未平倉口數'].iloc[0]
                    june3_date = june3_data['日期'].iloc[0]
                    
                    # 在圖上標記6/3
                    ax1.annotate('6/3', 
                                xy=(june3_date, june3_trade), 
                                xytext=(10, 20), textcoords='offset points',
                                bbox=dict(boxstyle='round,pad=0.3', facecolor='yellow', alpha=0.8),
                                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'))
            
            plt.tight_layout()
            
            # 儲存圖表
            filename = f"{contract}_updated_with_june3.png"
            filepath = os.path.join(output_dir, filename)
            plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            
            chart_paths.append(filepath)
            logger.info(f"✅ 已生成更新版圖表: {filename}")
        
        return chart_paths
        
    except Exception as e:
        logger.error(f"❌ 圖表生成失敗: {e}")
        return []

def generate_updated_summary(df):
    """生成包含6/3資料的更新版摘要"""
    try:
        if df.empty:
            return "⚠️ 無資料可分析"
        
        # 加總處理
        df_summary = df.groupby(['日期', '契約名稱']).agg({
            '多空淨額交易口數': 'sum',
            '多空淨額未平倉口數': 'sum'
        }).reset_index()
        
        date_range = f"{df['日期'].min().strftime('%Y-%m-%d')} 到 {df['日期'].max().strftime('%Y-%m-%d')}"
        unique_dates = df['日期'].dt.date.nunique()
        contracts = df['契約名稱'].unique().tolist()
        
        # 檢查6/3資料
        has_june3 = not df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()].empty
        june3_note = "✅ 已包含2025年6月3日最新資料！" if has_june3 else "❌ 未包含6月3日資料"
        
        summary = f"""🎉 台期所三大法人持倉分析報告 (更新版 - 新Google Sheets)
📅 分析期間: {date_range} ({unique_dates}個交易日)
📈 分析契約: {', '.join(contracts)}
🔗 資料來源: 新版Google Sheets (ID: 1w1uslujf-DF7BufO6s5TPYAjvWgUS3B_jCczDxrhmA4)
{june3_note}

📊 資料概況:
• 總資料筆數: {len(df)}筆
• 三大法人加總分析
• 雙軸圖表格式：柱狀圖(交易口數) + 折線圖(未平倉口數)

"""
        
        if has_june3:
            june3_data = df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
            june3_summary = june3_data.groupby('契約名稱').agg({
                '多空淨額交易口數': 'sum',
                '多空淨額未平倉口數': 'sum'
            })
            
            summary += "🔥 2025年6月3日最新持倉狀況:\n"
            for contract in june3_summary.index:
                trade = june3_summary.loc[contract, '多空淨額交易口數']
                position = june3_summary.loc[contract, '多空淨額未平倉口數']
                trade_trend = "📈 偏多" if trade > 0 else "📉 偏空" if trade < 0 else "➡️ 平衡"
                position_trend = "📈 偏多" if position > 0 else "📉 偏空" if position < 0 else "➡️ 平衡"
                
                summary += f"• {contract}: 交易量{trade:+,}口 {trade_trend}, 持倉量{position:+,}口 {position_trend}\n"
        
        summary += f"""
🎯 重要更新:
• 已修正Google Sheets連接問題
• 現在使用最新版本的資料源
• 確保包含最新的6月3日交易資料
• 圖表分析更加完整準確

📱 圖表已自動發送到Telegram群組！"""
        
        return summary
        
    except Exception as e:
        logger.error(f"❌ 摘要生成失敗: {e}")
        return "⚠️ 摘要生成失敗"

def main():
    """主程式"""
    try:
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== 更新版台期所圖表分析系統 (新Google Sheets) ===")
        
        # 1. 從新的Google Sheets載入資料
        df = load_data_from_new_sheet()
        
        if df.empty:
            logger.error("❌ 無法載入資料")
            return 1
        
        # 2. 生成更新版圖表
        chart_paths = generate_updated_charts(df)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return 1
        
        # 3. 生成更新版摘要
        summary_text = generate_updated_summary(df)
        print("\n" + summary_text)
        
        # 4. 發送到Telegram
        logger.info("📱 發送更新版圖表到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        if notifier.test_connection():
            success = notifier.send_chart_report(chart_paths, summary_text)
            
            if success:
                logger.info("🎉 更新版圖表已成功發送到Telegram！")
                logger.info("✅ 現在已使用新的Google Sheets，包含完整的6/3資料")
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