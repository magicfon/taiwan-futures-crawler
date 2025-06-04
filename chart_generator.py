#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
台期所資料圖表生成器
生成多空淨額交易與未平倉分析圖表
"""

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from pathlib import Path
import logging
import matplotlib as mpl
import platform
import json

# 根據作業系統設定中文字體
if platform.system() == 'Windows':
    # Windows 系統
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
elif platform.system() == 'Darwin':
    # macOS 系統
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Heiti TC', 'PingFang TC']
else:
    # Linux 系統
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'WenQuanYi Micro Hei', 'SimHei']

plt.rcParams['axes.unicode_minus'] = False

# 設定字體大小
mpl.rcParams.update({'font.size': 12})

logger = logging.getLogger("圖表生成器")

# 契約名稱對照表
CONTRACT_NAMES = {
    'TX': '臺股期貨',
    'TE': '電子期貨',
    'MTX': '小型臺指期貨',
    'ZMX': '微型臺指期貨',
    'NQF': '美國那斯達克100期貨'
}

class ChartGenerator:
    def __init__(self, output_dir="charts"):
        """
        初始化圖表生成器
        
        Args:
            output_dir: 圖表輸出目錄
        """
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 設定圖表樣式
        plt.style.use('seaborn-v0_8')  # 使用現代化樣式
        
    def prepare_data(self, df):
        """
        準備圖表數據
        
        Args:
            df: 包含期貨資料的DataFrame
            
        Returns:
            dict: 按契約分組的數據
        """
        if df.empty:
            logger.warning("沒有資料可以繪製圖表")
            return {}
        
        # 確保有必要的欄位
        required_columns = ['日期', '契約名稱']
        if not all(col in df.columns for col in required_columns):
            logger.error(f"資料缺少必要欄位: {required_columns}")
            return {}
        
        # 轉換日期欄位
        if '日期' in df.columns:
            df['日期'] = pd.to_datetime(df['日期'])
        
        # 按契約分組
        contract_data = {}
        for contract in df['契約名稱'].unique():
            if pd.isna(contract):
                continue
                
            contract_df = df[df['契約名稱'] == contract].copy()
            contract_df = contract_df.sort_values('日期')
            
            # 如果有身份別資料，合併計算總計
            if '身份別' in df.columns:
                # 按日期分組，計算各身份別的總和
                daily_summary = contract_df.groupby('日期').agg({
                    '多空淨額交易口數': 'sum',
                    '多空淨額未平倉口數': 'sum'
                }).reset_index()
            else:
                daily_summary = contract_df[['日期', '多空淨額交易口數', '多空淨額未平倉口數']].copy()
            
            contract_data[contract] = daily_summary
        
        return contract_data
    
    def create_dual_axis_chart(self, contract, data, save_path=None):
        """
        創建雙軸圖表
        
        Args:
            contract: 契約代碼
            data: 該契約的資料DataFrame
            save_path: 儲存路徑
            
        Returns:
            str: 圖表檔案路徑
        """
        if data.empty:
            logger.warning(f"契約 {contract} 沒有資料可繪製")
            return None
        
        # 創建圖表
        fig, ax1 = plt.subplots(figsize=(14, 8))
        
        # 獲取契約全名
        contract_name = CONTRACT_NAMES.get(contract, contract)
        
        # 主軸 - 多空淨額交易口數
        color1 = '#1f77b4'  # 藍色
        ax1.set_xlabel('日期', fontsize=12)
        ax1.set_ylabel('多空淨額交易口數', color=color1, fontsize=12)
        
        # 繪製交易口數柱狀圖
        bars = ax1.bar(data['日期'], data['多空淨額交易口數'], 
                      color=color1, alpha=0.6, label='多空淨額交易口數', width=0.8)
        
        # 在柱狀圖上添加正負數的顏色區分
        for i, bar in enumerate(bars):
            if data.iloc[i]['多空淨額交易口數'] >= 0:
                bar.set_color('#2ca02c')  # 正數用綠色
            else:
                bar.set_color('#d62728')  # 負數用紅色
        
        ax1.tick_params(axis='y', labelcolor=color1)
        ax1.grid(True, alpha=0.3)
        
        # 副軸 - 多空淨額未平倉口數
        ax2 = ax1.twinx()
        color2 = '#ff7f0e'  # 橙色
        ax2.set_ylabel('多空淨額未平倉口數', color=color2, fontsize=12)
        
        # 繪製未平倉口數線圖
        line = ax2.plot(data['日期'], data['多空淨額未平倉口數'], 
                       color=color2, linewidth=3, marker='o', markersize=4,
                       label='多空淨額未平倉口數')
        
        ax2.tick_params(axis='y', labelcolor=color2)
        
        # 設定標題
        title = f"{contract_name} ({contract}) - 30天持倉分析"
        plt.title(title, fontsize=16, fontweight='bold', pad=20)
        
        # 設定x軸日期格式
        ax1.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(data)//10)))
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
        
        # 添加零線參考
        ax1.axhline(y=0, color='black', linestyle='-', alpha=0.3, linewidth=1)
        ax2.axhline(y=0, color='black', linestyle='--', alpha=0.3, linewidth=1)
        
        # 添加圖例
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, 
                  loc='upper left', bbox_to_anchor=(0.02, 0.98))
        
        # 添加數據摘要
        latest_trade = data.iloc[-1]['多空淨額交易口數'] if not data.empty else 0
        latest_position = data.iloc[-1]['多空淨額未平倉口數'] if not data.empty else 0
        latest_date = data.iloc[-1]['日期'].strftime('%Y-%m-%d') if not data.empty else 'N/A'
        
        # 計算趨勢
        if len(data) >= 2:
            trade_trend = "↗" if data.iloc[-1]['多空淨額交易口數'] > data.iloc[-2]['多空淨額交易口數'] else "↘"
            position_trend = "↗" if data.iloc[-1]['多空淨額未平倉口數'] > data.iloc[-2]['多空淨額未平倉口數'] else "↘"
        else:
            trade_trend = position_trend = "-"
        
        summary_text = f"最新資料 ({latest_date})\n"
        summary_text += f"交易口數: {latest_trade:,} {trade_trend}\n"
        summary_text += f"未平倉口數: {latest_position:,} {position_trend}"
        
        # 在右上角添加摘要文字
        ax1.text(0.98, 0.98, summary_text, transform=ax1.transAxes,
                verticalalignment='top', horizontalalignment='right',
                bbox=dict(boxstyle='round', facecolor='white', alpha=0.8),
                fontsize=10)
        
        # 調整佈局
        plt.tight_layout()
        
        # 儲存圖表
        if not save_path:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{contract}_{contract_name}_持倉分析_{timestamp}.png"
            save_path = os.path.join(self.output_dir, filename)
        
        try:
            plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            logger.info(f"📊 圖表已儲存: {save_path}")
            return save_path
        except Exception as e:
            logger.error(f"儲存圖表失敗: {e}")
            plt.close()
            return None
    
    def create_overview_chart(self, contract_data, save_path=None):
        """
        創建所有契約的概覽圖表
        
        Args:
            contract_data: 所有契約的資料字典
            save_path: 儲存路徑
            
        Returns:
            str: 圖表檔案路徑
        """
        if not contract_data:
            logger.warning("沒有資料可繪製概覽圖表")
            return None
        
        # 計算子圖配置
        n_contracts = len(contract_data)
        if n_contracts <= 2:
            rows, cols = 1, n_contracts
            figsize = (7*n_contracts, 6)
        elif n_contracts <= 4:
            rows, cols = 2, 2
            figsize = (14, 12)
        else:
            rows, cols = 3, 2
            figsize = (14, 18)
        
        fig, axes = plt.subplots(rows, cols, figsize=figsize)
        if n_contracts == 1:
            axes = [axes]
        elif rows == 1 or cols == 1:
            axes = axes.flatten()
        else:
            axes = axes.flatten()
        
        # 為每個契約繪製子圖
        for i, (contract, data) in enumerate(contract_data.items()):
            if i >= len(axes):
                break
                
            ax = axes[i]
            contract_name = CONTRACT_NAMES.get(contract, contract)
            
            if data.empty:
                ax.text(0.5, 0.5, f'{contract_name}\n無資料', 
                       ha='center', va='center', transform=ax.transAxes)
                ax.set_title(f"{contract_name} ({contract})")
                continue
            
            # 繪製交易口數柱狀圖
            bars = ax.bar(data['日期'], data['多空淨額交易口數'], 
                         alpha=0.7, width=0.8)
            
            # 設定柱狀圖顏色
            for j, bar in enumerate(bars):
                if data.iloc[j]['多空淨額交易口數'] >= 0:
                    bar.set_color('#2ca02c')
                else:
                    bar.set_color('#d62728')
            
            # 添加未平倉口數線圖（副軸）
            ax2 = ax.twinx()
            ax2.plot(data['日期'], data['多空淨額未平倉口數'], 
                    color='#ff7f0e', linewidth=2, marker='o', markersize=3)
            
            # 設定標題和軸標籤
            ax.set_title(f"{contract_name} ({contract})", fontweight='bold')
            ax.set_ylabel('交易口數', fontsize=10)
            ax2.set_ylabel('未平倉口數', fontsize=10, color='#ff7f0e')
            
            # 設定日期格式
            ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(data)//5)))
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)
            
            # 添加零線
            ax.axhline(y=0, color='black', linestyle='-', alpha=0.3, linewidth=1)
            ax2.axhline(y=0, color='black', linestyle='--', alpha=0.3, linewidth=1)
            
            ax.grid(True, alpha=0.3)
        
        # 隱藏多餘的子圖
        for i in range(len(contract_data), len(axes)):
            axes[i].set_visible(False)
        
        # 設定主標題
        fig.suptitle('台期所30天持倉分析總覽', fontsize=20, fontweight='bold', y=0.98)
        
        # 調整佈局
        plt.tight_layout()
        plt.subplots_adjust(top=0.94)
        
        # 儲存圖表
        if not save_path:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"台期所持倉分析總覽_{timestamp}.png"
            save_path = os.path.join(self.output_dir, filename)
        
        try:
            plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            logger.info(f"📊 概覽圖表已儲存: {save_path}")
            return save_path
        except Exception as e:
            logger.error(f"儲存概覽圖表失敗: {e}")
            plt.close()
            return None
    
    def generate_all_charts(self, df):
        """
        生成所有圖表
        
        Args:
            df: 包含期貨資料的DataFrame
            
        Returns:
            list: 生成的圖表檔案路徑列表
        """
        chart_paths = []
        
        # 準備數據
        contract_data = self.prepare_data(df)
        
        if not contract_data:
            logger.warning("沒有有效數據可生成圖表")
            return chart_paths
        
        # 生成概覽圖表
        overview_path = self.create_overview_chart(contract_data)
        if overview_path:
            chart_paths.append(overview_path)
        
        # 為每個契約生成獨立圖表
        for contract, data in contract_data.items():
            chart_path = self.create_dual_axis_chart(contract, data)
            if chart_path:
                chart_paths.append(chart_path)
        
        logger.info(f"已生成 {len(chart_paths)} 個圖表")
        return chart_paths
    
    def generate_summary_text(self, df):
        """
        生成摘要文字
        
        Args:
            df: 包含期貨資料的DataFrame
            
        Returns:
            str: 摘要文字
        """
        if df.empty:
            return "📊 *台期所持倉分析報告*\n\n❌ 今日無交易資料"
        
        # 獲取最新日期
        latest_date = df['日期'].max().strftime('%Y-%m-%d') if '日期' in df.columns else datetime.now().strftime('%Y-%m-%d')
        
        # 統計契約數量
        unique_contracts = df['契約名稱'].nunique() if '契約名稱' in df.columns else 0
        
        # 生成摘要
        summary = f"📊 *台期所持倉分析報告*\n\n"
        summary += f"📅 資料日期: {latest_date}\n"
        summary += f"📈 契約數量: {unique_contracts} 個\n"
        summary += f"📊 分析期間: 最近30個交易日\n\n"
        
        # 各契約簡要說明
        if '契約名稱' in df.columns:
            summary += "*各契約持倉狀況:*\n"
            for contract in df['契約名稱'].unique():
                if pd.isna(contract):
                    continue
                contract_name = CONTRACT_NAMES.get(contract, contract)
                contract_data = df[df['契約名稱'] == contract]
                
                # 計算最新的淨額
                if not contract_data.empty:
                    latest_trade = contract_data['多空淨額交易口數'].iloc[-1] if '多空淨額交易口數' in contract_data.columns else 0
                    latest_position = contract_data['多空淨額未平倉口數'].iloc[-1] if '多空淨額未平倉口數' in contract_data.columns else 0
                    
                    # 判斷多空傾向
                    trade_bias = "偏多" if latest_trade > 0 else "偏空" if latest_trade < 0 else "中性"
                    position_bias = "偏多" if latest_position > 0 else "偏空" if latest_position < 0 else "中性"
                    
                    summary += f"• {contract_name}: 交易{trade_bias}, 持倉{position_bias}\n"
        
        summary += f"\n⏰ 更新時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        return summary

    def load_data_from_google_sheets(self, days=30):
        """
        從Google Sheets載入指定天數的歷史資料
        
        Args:
            days: 要載入的天數
            
        Returns:
            DataFrame: 包含歷史資料的DataFrame
        """
        try:
            from google_sheets_manager import GoogleSheetsManager
            
            # 初始化Google Sheets管理器
            sheets_manager = GoogleSheetsManager()
            
            if not sheets_manager.client:
                logger.warning("Google Sheets未啟用，無法載入歷史資料")
                return pd.DataFrame()
            
            # 嘗試載入現有試算表
            config_file = Path("config/spreadsheet_config.json")
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    spreadsheet_id = config.get('spreadsheet_id')
                    
                    if spreadsheet_id:
                        sheets_manager.connect_spreadsheet(spreadsheet_id)
                        logger.info("已連接到Google試算表")
            
            if not sheets_manager.spreadsheet:
                logger.warning("無法連接到Google試算表")
                return pd.DataFrame()
            
            # 從「原始資料」工作表讀取資料
            try:
                worksheet = sheets_manager.spreadsheet.worksheet("原始資料")
                data = worksheet.get_all_records()
                
                if not data:
                    logger.warning("Google Sheets中沒有找到資料")
                    return pd.DataFrame()
                
                df = pd.DataFrame(data)
                
                # 資料清理和格式轉換
                if '日期' in df.columns:
                    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                    
                    # 過濾最近N天的資料
                    end_date = datetime.now()
                    start_date = end_date - timedelta(days=days)
                    df = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
                    
                    # 只保留工作日
                    df = df[df['日期'].dt.weekday < 5]
                    
                    logger.info(f"從Google Sheets載入了 {len(df)} 筆近{days}天資料")
                    logger.info(f"日期範圍: {df['日期'].min()} 到 {df['日期'].max()}")
                    
                    return df
                
            except Exception as e:
                logger.warning(f"無法從「原始資料」工作表讀取: {e}")
                
                # 嘗試從其他可能的工作表讀取
                worksheets = sheets_manager.spreadsheet.worksheets()
                for ws in worksheets:
                    if '資料' in ws.title or 'data' in ws.title.lower():
                        try:
                            data = ws.get_all_records()
                            if data:
                                df = pd.DataFrame(data)
                                if '日期' in df.columns:
                                    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
                                    end_date = datetime.now()
                                    start_date = end_date - timedelta(days=days)
                                    df = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
                                    df = df[df['日期'].dt.weekday < 5]
                                    
                                    logger.info(f"從工作表「{ws.title}」載入了 {len(df)} 筆資料")
                                    return df
                        except:
                            continue
            
            logger.warning("無法從Google Sheets載入有效資料")
            return pd.DataFrame()
            
        except ImportError:
            logger.warning("Google Sheets模組未安裝，無法載入歷史資料")
            return pd.DataFrame()
        except Exception as e:
            logger.error(f"從Google Sheets載入資料時發生錯誤: {e}")
            return pd.DataFrame()

# 測試函數
def test_chart_generator():
    """測試圖表生成功能"""
    # 創建測試數據
    dates = pd.date_range(start='2024-01-01', periods=30, freq='D')
    
    test_data = []
    for date in dates:
        for contract in ['TX', 'TE']:
            test_data.append({
                '日期': date,
                '契約名稱': contract,
                '多空淨額交易口數': np.random.randint(-5000, 5000),
                '多空淨額未平倉口數': np.random.randint(-10000, 10000)
            })
    
    df = pd.DataFrame(test_data)
    
    # 生成圖表
    generator = ChartGenerator()
    chart_paths = generator.generate_all_charts(df)
    summary_text = generator.generate_summary_text(df)
    
    print("生成的圖表:")
    for path in chart_paths:
        print(f"  {path}")
    
    print(f"\n摘要文字:\n{summary_text}")

if __name__ == "__main__":
    test_chart_generator() 