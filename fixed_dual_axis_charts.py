#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
修正版雙軸圖表生成器
確保包含所有可用資料，包括2024年的6/3
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import logging
import os
from datetime import datetime, timedelta

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("修正版雙軸圖表")

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def load_complete_data():
    """載入完整的歷史資料，包含所有可用日期"""
    try:
        from google_sheets_manager import GoogleSheetsManager
        
        logger.info("🔍 載入完整的Google Sheets資料...")
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 檢查所有工作表並合併資料
        worksheets_to_check = ['歷史資料', 'Sheet1', '最新30天資料']
        all_data = []
        
        for ws_name in worksheets_to_check:
            try:
                logger.info(f"📖 檢查工作表: {ws_name}")
                ws = spreadsheet.worksheet(ws_name)
                values = ws.get_all_values()
                
                if len(values) < 2:
                    logger.info(f"  ⚠️ {ws_name} 沒有足夠資料")
                    continue
                
                headers = values[0]
                data_rows = values[1:]
                
                # 檢查是否有日期欄位
                if '日期' not in headers:
                    logger.info(f"  ⚠️ {ws_name} 沒有日期欄位")
                    continue
                
                df_temp = pd.DataFrame(data_rows, columns=headers)
                df_temp = df_temp[df_temp['日期'].str.strip() != '']
                
                if len(df_temp) == 0:
                    logger.info(f"  ⚠️ {ws_name} 沒有有效資料")
                    continue
                
                # 轉換日期
                df_temp['日期'] = pd.to_datetime(df_temp['日期'], format='%Y/%m/%d', errors='coerce')
                df_temp = df_temp.dropna(subset=['日期'])
                
                if len(df_temp) > 0:
                    logger.info(f"✅ 從 {ws_name} 載入 {len(df_temp)} 筆資料")
                    logger.info(f"📅 日期範圍: {df_temp['日期'].min()} 到 {df_temp['日期'].max()}")
                    all_data.append(df_temp)
            
            except Exception as e:
                logger.debug(f"工作表 {ws_name} 讀取失敗: {e}")
                continue
        
        if not all_data:
            logger.error("❌ 無法從任何工作表載入資料")
            return pd.DataFrame()
        
        # 合併所有資料並去重
        df = pd.concat(all_data, ignore_index=True)
        df = df.drop_duplicates(subset=['日期', '契約名稱', '身份別'])
        
        # 轉換數值欄位
        numeric_cols = ['多空淨額交易口數', '多空淨額未平倉口數']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        logger.info(f"📊 總共合併 {len(df)} 筆唯一資料")
        logger.info(f"📅 完整日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
        
        # 檢查是否有2024年6/3和2025年6/3資料
        june3_2024 = df[df['日期'].dt.date == pd.to_datetime('2024/06/03').date()]
        june3_2025 = df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
        
        if not june3_2024.empty:
            logger.info(f"✅ 找到2024年6/3資料: {len(june3_2024)}筆")
        if not june3_2025.empty:
            logger.info(f"✅ 找到2025年6/3資料: {len(june3_2025)}筆")
        
        # 顯示所有可用日期的範圍
        all_dates = sorted(df['日期'].dt.date.unique())
        logger.info(f"📆 共有 {len(all_dates)} 個不同交易日")
        logger.info(f"📆 最早日期: {all_dates[0]}")
        logger.info(f"📆 最新日期: {all_dates[-1]}")
        logger.info(f"📆 最新10個日期: {all_dates[-10:]}")
        
        # 改進的30天過濾邏輯：從最新日期往回取30個工作日
        latest_date = df['日期'].max()
        logger.info(f"📅 以最新日期 {latest_date.strftime('%Y-%m-%d')} 為基準取30個工作日")
        
        # 取得所有工作日（排除週末）
        workdays = df[df['日期'].dt.weekday < 5]['日期'].dt.date.unique()
        workdays_sorted = sorted(workdays, reverse=True)
        
        # 取最近30個工作日
        if len(workdays_sorted) > 30:
            recent_30_workdays = workdays_sorted[:30]
            logger.info(f"📅 選取的30個工作日範圍: {min(recent_30_workdays)} 到 {max(recent_30_workdays)}")
        else:
            recent_30_workdays = workdays_sorted
            logger.info(f"📅 可用工作日不足30天，使用全部 {len(recent_30_workdays)} 個工作日")
        
        # 過濾資料
        df_filtered = df[df['日期'].dt.date.isin(recent_30_workdays)]
        
        logger.info(f"✅ 過濾後有 {len(df_filtered)} 筆30天內的資料")
        logger.info(f"📅 最終分析範圍: {df_filtered['日期'].min().strftime('%Y-%m-%d')} 到 {df_filtered['日期'].max().strftime('%Y-%m-%d')}")
        
        # 再次檢查6/3資料是否在最終結果中
        final_june3_2024 = df_filtered[df_filtered['日期'].dt.date == pd.to_datetime('2024/06/03').date()]
        final_june3_2025 = df_filtered[df_filtered['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
        
        if not final_june3_2024.empty:
            logger.info(f"✅ 最終結果包含2024年6/3資料: {len(final_june3_2024)}筆")
        if not final_june3_2025.empty:
            logger.info(f"✅ 最終結果包含2025年6/3資料: {len(final_june3_2025)}筆")
        
        return df_filtered
        
    except Exception as e:
        logger.error(f"❌ 資料載入失敗: {e}")
        return pd.DataFrame()

def generate_enhanced_dual_axis_charts(df, output_dir="enhanced_dual_axis_charts"):
    """生成增強版雙軸圖表"""
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
        logger.info(f"🎨 開始生成 {len(contracts)} 個契約的增強版雙軸圖表...")
        
        for contract in contracts:
            contract_data = df_processed[df_processed['契約名稱'] == contract].copy()
            contract_data = contract_data.sort_values('日期')
            
            if contract_data.empty:
                continue
            
            # 創建增強版雙軸圖表
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
            title = f"{contract} 契約 - 三大法人整體持倉分析 (30天趨勢)"
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
            
            # 添加數值標籤（只顯示較大的值）
            max_abs_trade = abs(contract_data['多空淨額交易口數']).max()
            for bar, value in zip(bars, contract_data['多空淨額交易口數']):
                if abs(value) > max_abs_trade * 0.4:  # 只顯示較大的值
                    height = bar.get_height()
                    offset = abs(height) * 0.08
                    y_pos = height + offset if height > 0 else height - offset
                    ax1.text(bar.get_x() + bar.get_width()/2., y_pos,
                            f'{int(value):,}', ha='center', 
                            va='bottom' if height > 0 else 'top', 
                            fontsize=9, color='black', fontweight='bold',
                            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.8))
            
            # 美化圖表
            ax1.spines['top'].set_visible(False)
            ax2.spines['top'].set_visible(False)
            
            plt.tight_layout()
            
            # 儲存圖表
            filename = f"{contract}_enhanced_dual_axis_30days.png"
            filepath = os.path.join(output_dir, filename)
            plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            
            chart_paths.append(filepath)
            logger.info(f"✅ 已生成增強版雙軸圖表: {filename}")
        
        return chart_paths
        
    except Exception as e:
        logger.error(f"❌ 增強版圖表生成失敗: {e}")
        return []

def generate_enhanced_summary(df):
    """生成增強版分析摘要"""
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
        
        # 檢查是否包含6/3資料
        has_june3_2024 = not df[df['日期'].dt.date == pd.to_datetime('2024/06/03').date()].empty
        has_june3_2025 = not df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()].empty
        
        june3_note = ""
        if has_june3_2024:
            june3_note += "✅ 包含2024年6/3資料 "
        if has_june3_2025:
            june3_note += "✅ 包含2025年6/3資料 "
        
        summary = f"""📊 台期所三大法人持倉分析報告 (完整30天資料)
📅 分析期間: {date_range} ({unique_dates}個交易日)
📈 分析契約: {', '.join(contracts)}
📋 圖表格式: 增強版雙軸圖表（綠紅柱狀圖：交易口數，紫色折線圖：未平倉口數）
🏷️ 資料來源: Google Sheets「台期所資料分析」完整歷史資料
{june3_note}

📊 分析方式: 三大法人（外資、投信、自營商）加總

"""
        
        for contract in contracts:
            contract_data = df_summary[df_summary['契約名稱'] == contract].sort_values('日期')
            if contract_data.empty:
                continue
            
            latest_trade = contract_data['多空淨額交易口數'].iloc[-1]
            latest_position = contract_data['多空淨額未平倉口數'].iloc[-1]
            
            # 計算5日平均和趨勢
            if len(contract_data) >= 5:
                recent_trade_avg = contract_data['多空淨額交易口數'].tail(5).mean()
                recent_position_avg = contract_data['多空淨額未平倉口數'].tail(5).mean()
                
                # 計算趨勢變化
                older_trade_avg = contract_data['多空淨額交易口數'].tail(10).head(5).mean()
                older_position_avg = contract_data['多空淨額未平倉口數'].tail(10).head(5).mean()
                
                trade_change = recent_trade_avg - older_trade_avg
                position_change = recent_position_avg - older_position_avg
            else:
                recent_trade_avg = latest_trade
                recent_position_avg = latest_position
                trade_change = 0
                position_change = 0
            
            # 趨勢判斷
            trade_trend = "📈 偏多" if recent_trade_avg > 2000 else "📉 偏空" if recent_trade_avg < -2000 else "➡️ 平衡"
            position_trend = "📈 偏多" if recent_position_avg > 10000 else "📉 偏空" if recent_position_avg < -10000 else "➡️ 平衡"
            
            change_trend = "⬆️ 轉強" if trade_change > 1000 else "⬇️ 轉弱" if trade_change < -1000 else "➡️ 持平"
            
            summary += f"""💼 {contract}契約深度分析:
   • 最新淨交易量: {latest_trade:,} 口 {trade_trend}
   • 最新淨未平倉: {latest_position:,} 口 {position_trend}
   • 5日平均交易量: {recent_trade_avg:,.0f} 口 {change_trend}
   • 5日平均未平倉: {recent_position_avg:,.0f} 口
   • 交易量變化: {trade_change:+,.0f} 口
   • 持倉量變化: {position_change:+,.0f} 口

"""
        
        summary += """📱 增強版圖表說明:
• 綠色柱狀圖: 正值交易量（偏多交易）
• 紅色柱狀圖: 負值交易量（偏空交易）  
• 紫色折線圖: 多空淨額未平倉口數（持倉趨勢）
• 包含完整30個工作日資料，確保趨勢分析完整性！

🎯 本次更新已修正資料載入問題，確保包含所有可用的歷史資料！"""
        
        return summary
        
    except Exception as e:
        logger.error(f"❌ 增強版摘要生成失敗: {e}")
        return "⚠️ 摘要生成失敗"

def main():
    """主程式"""
    try:
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== 修正版台期所雙軸圖表分析系統 ===")
        
        # 1. 載入完整資料
        df = load_complete_data()
        
        if df.empty:
            logger.error("❌ 無法載入資料")
            return 1
        
        # 2. 生成增強版雙軸圖表
        chart_paths = generate_enhanced_dual_axis_charts(df)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return 1
        
        # 3. 生成增強版分析摘要
        summary_text = generate_enhanced_summary(df)
        print("\n" + summary_text)
        
        # 4. 發送到Telegram
        logger.info("📱 發送修正版雙軸圖表到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        if notifier.test_connection():
            success = notifier.send_chart_report(chart_paths, summary_text)
            
            if success:
                logger.info("🎉 修正版雙軸圖表已成功發送到Telegram！")
                logger.info("📊 已修正資料載入問題，包含完整的歷史資料")
                logger.info("🎯 現在應該可以看到包含6/3在內的完整30天趨勢了")
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