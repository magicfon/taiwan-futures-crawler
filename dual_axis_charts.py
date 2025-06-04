#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
雙軸圖表生成器
主軸：多空淨額交易口數（柱狀圖）
副軸：多空淨額未平倉口數（折線圖）
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import logging
import os
from datetime import datetime, timedelta

# 設定日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("雙軸圖表生成")

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def load_latest_data_with_june3():
    """載入包含6/3的最新資料"""
    try:
        from google_sheets_manager import GoogleSheetsManager
        
        logger.info("🔍 載入最新Google Sheets資料（包含6/3）...")
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 檢查多個可能的工作表
        worksheets_to_check = ['歷史資料', 'Sheet1', '最新30天資料']
        
        all_data = []
        
        for ws_name in worksheets_to_check:
            try:
                logger.info(f"📖 檢查工作表: {ws_name}")
                ws = spreadsheet.worksheet(ws_name)
                values = ws.get_all_values()
                
                if len(values) < 2:
                    continue
                
                headers = values[0]
                data_rows = values[1:]
                
                df_temp = pd.DataFrame(data_rows, columns=headers)
                df_temp = df_temp[df_temp['日期'].str.strip() != '']
                
                if len(df_temp) > 0:
                    # 嘗試轉換日期
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
        logger.info(f"📅 最終日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
        
        # 檢查是否有6/3資料
        june_3_data = df[df['日期'].dt.date == pd.to_datetime('2025/06/03').date()]
        if not june_3_data.empty:
            logger.info(f"✅ 找到6/3資料: {len(june_3_data)}筆")
        else:
            logger.warning("⚠️ 沒有找到6/3資料，檢查今天的日期或工作表內容")
            
            # 顯示最新幾天的資料
            latest_dates = sorted(df['日期'].dt.date.unique(), reverse=True)
            logger.info(f"📆 最新5個日期: {latest_dates[:5]}")
        
        # 過濾最近30個工作日
        end_date = df['日期'].max()  # 使用資料中的最新日期，而不是今天
        start_date = end_date - timedelta(days=45)
        
        df_filtered = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
        df_filtered = df_filtered[df_filtered['日期'].dt.weekday < 5]
        
        # 取最近30個工作日
        unique_dates = sorted(df_filtered['日期'].dt.date.unique())
        if len(unique_dates) > 30:
            recent_30_dates = unique_dates[-30:]
            df_filtered = df_filtered[df_filtered['日期'].dt.date.isin(recent_30_dates)]
        
        logger.info(f"✅ 過濾後有 {len(df_filtered)} 筆30天內的資料")
        logger.info(f"📅 分析範圍: {df_filtered['日期'].min().strftime('%Y-%m-%d')} 到 {df_filtered['日期'].max().strftime('%Y-%m-%d')}")
        
        return df_filtered
        
    except Exception as e:
        logger.error(f"❌ 資料載入失敗: {e}")
        return pd.DataFrame()

def generate_dual_axis_charts(df, output_dir="dual_axis_charts", aggregate_identities=True):
    """
    生成雙軸圖表
    
    Args:
        df: 資料DataFrame
        output_dir: 輸出目錄
        aggregate_identities: 是否加總三大法人
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        chart_paths = []
        
        # 處理資料
        if aggregate_identities:
            logger.info("📊 加總三大法人資料...")
            df_processed = df.groupby(['日期', '契約名稱']).agg({
                '多空淨額交易口數': 'sum',
                '多空淨額未平倉口數': 'sum'
            }).reset_index()
        else:
            df_processed = df.copy()
        
        contracts = df_processed['契約名稱'].unique()
        logger.info(f"🎨 開始生成 {len(contracts)} 個契約的雙軸圖表...")
        
        for contract in contracts:
            contract_data = df_processed[df_processed['契約名稱'] == contract].copy()
            contract_data = contract_data.sort_values('日期')
            
            if contract_data.empty:
                continue
            
            # 創建雙軸圖表
            fig, ax1 = plt.subplots(figsize=(15, 8))
            
            # 主軸：交易口數（柱狀圖）
            bars = ax1.bar(contract_data['日期'], contract_data['多空淨額交易口數'], 
                          alpha=0.7, color='#2E86AB', width=0.8, label='多空淨額交易口數')
            
            # 設定主軸
            ax1.set_xlabel('日期', fontsize=12, fontweight='bold')
            ax1.set_ylabel('多空淨額交易口數', fontsize=12, fontweight='bold', color='#2E86AB')
            ax1.tick_params(axis='y', labelcolor='#2E86AB')
            ax1.grid(True, alpha=0.3)
            ax1.axhline(y=0, color='gray', linestyle='-', alpha=0.5, linewidth=1)
            
            # 副軸：未平倉口數（折線圖）
            ax2 = ax1.twinx()
            line = ax2.plot(contract_data['日期'], contract_data['多空淨額未平倉口數'], 
                           color='#E74C3C', linewidth=3, marker='o', markersize=6, 
                           label='多空淨額未平倉口數')
            
            # 設定副軸
            ax2.set_ylabel('多空淨額未平倉口數', fontsize=12, fontweight='bold', color='#E74C3C')
            ax2.tick_params(axis='y', labelcolor='#E74C3C')
            ax2.axhline(y=0, color='gray', linestyle='--', alpha=0.5, linewidth=1)
            
            # 標題和格式
            title = f"{contract} 契約 - 三大法人{'整體' if aggregate_identities else '分別'}持倉分析 (30天)"
            plt.title(title, fontsize=16, fontweight='bold', pad=20)
            
            # 設定日期格式
            ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
            ax1.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
            plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
            
            # 添加圖例
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=10)
            
            # 添加數值標籤到柱狀圖（僅顯示絕對值較大的）
            for bar, value in zip(bars, contract_data['多空淨額交易口數']):
                if abs(value) > abs(contract_data['多空淨額交易口數']).max() * 0.3:  # 只顯示較大的值
                    height = bar.get_height()
                    ax1.text(bar.get_x() + bar.get_width()/2., height + (abs(height) * 0.05 if height > 0 else -abs(height) * 0.05),
                            f'{int(value):,}', ha='center', va='bottom' if height > 0 else 'top', 
                            fontsize=8, color='#2E86AB', fontweight='bold')
            
            plt.tight_layout()
            
            # 儲存圖表
            filename = f"{contract}_dual_axis_{'aggregated' if aggregate_identities else 'separate'}_30days.png"
            filepath = os.path.join(output_dir, filename)
            plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            
            chart_paths.append(filepath)
            logger.info(f"✅ 已生成雙軸圖表: {filename}")
        
        return chart_paths
        
    except Exception as e:
        logger.error(f"❌ 雙軸圖表生成失敗: {e}")
        return []

def generate_report_summary(df, is_aggregated=True):
    """生成分析報告摘要"""
    try:
        if df.empty:
            return "⚠️ 無資料可分析"
        
        date_range = f"{df['日期'].min().strftime('%Y-%m-%d')} 到 {df['日期'].max().strftime('%Y-%m-%d')}"
        unique_dates = df['日期'].dt.date.nunique()
        contracts = df['契約名稱'].unique().tolist()
        
        # 如果是加總模式，先處理資料
        if is_aggregated:
            df_summary = df.groupby(['日期', '契約名稱']).agg({
                '多空淨額交易口數': 'sum',
                '多空淨額未平倉口數': 'sum'
            }).reset_index()
        else:
            df_summary = df.copy()
        
        summary = f"""📊 台期所三大法人持倉分析報告 (30天)
📅 分析期間: {date_range} ({unique_dates}個交易日)
📈 分析契約: {', '.join(contracts)}
📋 圖表格式: 雙軸圖表（主軸柱狀圖：交易口數，副軸折線圖：未平倉口數）
🏷️ 資料來源: Google Sheets「台期所資料分析」

"""
        
        if is_aggregated:
            summary += "📊 分析方式: 三大法人（外資、投信、自營商）加總\n\n"
            
            for contract in contracts:
                contract_data = df_summary[df_summary['契約名稱'] == contract].sort_values('日期')
                if contract_data.empty:
                    continue
                
                latest_trade = contract_data['多空淨額交易口數'].iloc[-1]
                latest_position = contract_data['多空淨額未平倉口數'].iloc[-1]
                
                # 計算趨勢
                if len(contract_data) >= 5:
                    recent_trade_avg = contract_data['多空淨額交易口數'].tail(5).mean()
                    recent_position_avg = contract_data['多空淨額未平倉口數'].tail(5).mean()
                else:
                    recent_trade_avg = latest_trade
                    recent_position_avg = latest_position
                
                trade_trend = "📈 偏多" if recent_trade_avg > 1000 else "📉 偏空" if recent_trade_avg < -1000 else "➡️ 平衡"
                position_trend = "📈 偏多" if recent_position_avg > 5000 else "📉 偏空" if recent_position_avg < -5000 else "➡️ 平衡"
                
                summary += f"""💼 {contract}契約分析:
   • 最新淨交易量: {latest_trade:,} 口 {trade_trend}
   • 最新淨未平倉: {latest_position:,} 口 {position_trend}
   • 5日平均交易量: {recent_trade_avg:,.0f} 口
   • 5日平均未平倉: {recent_position_avg:,.0f} 口

"""
        
        summary += """📱 圖表說明:
• 藍色柱狀圖（主軸）: 多空淨額交易口數，正值表示偏多交易，負值表示偏空交易
• 紅色折線圖（副軸）: 多空淨額未平倉口數，正值表示偏多持倉，負值表示偏空持倉
• 30天完整趨勢分析，助您掌握三大法人的市場動向！"""
        
        return summary
        
    except Exception as e:
        logger.error(f"❌ 摘要生成失敗: {e}")
        return "⚠️ 摘要生成失敗"

def main():
    """主程式"""
    try:
        from telegram_notifier import TelegramNotifier
        
        logger.info("=== 台期所雙軸圖表分析系統 ===")
        
        # 1. 載入最新資料
        df = load_latest_data_with_june3()
        
        if df.empty:
            logger.error("❌ 無法載入資料")
            return 1
        
        # 2. 生成雙軸圖表（加總模式）
        chart_paths = generate_dual_axis_charts(df, aggregate_identities=True)
        
        if not chart_paths:
            logger.error("❌ 圖表生成失敗")
            return 1
        
        # 3. 生成分析摘要
        summary_text = generate_report_summary(df, is_aggregated=True)
        print("\n" + summary_text)
        
        # 4. 發送到Telegram
        logger.info("📱 發送雙軸圖表到Telegram...")
        
        bot_token = "7088578241:AAErbP-EuoRGClRZ3FFfPMjl8k3CFpqgn8E"
        chat_id = "1038401606"
        notifier = TelegramNotifier(bot_token, chat_id)
        
        if notifier.test_connection():
            success = notifier.send_chart_report(chart_paths, summary_text)
            
            if success:
                logger.info("🎉 雙軸圖表已成功發送到Telegram！")
                logger.info("📊 新格式: 主軸柱狀圖（交易量）+ 副軸折線圖（未平倉）")
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