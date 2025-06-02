#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
台期所30天日報生成器
用於產生分析報告和派報
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import json
from pathlib import Path
import logging
from database_manager import TaifexDatabaseManager
import matplotlib.font_manager as fm

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Arial Unicode MS', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False

class DailyReportGenerator:
    """30天日報生成器"""
    
    def __init__(self, db_manager=None):
        """初始化日報生成器"""
        self.db_manager = db_manager or TaifexDatabaseManager()
        self.logger = logging.getLogger(__name__)
        self.output_dir = Path("reports")
        self.output_dir.mkdir(exist_ok=True)
    
    def generate_30day_report(self):
        """生成30天完整報告"""
        timestamp = datetime.now().strftime('%Y%m%d')
        
        # 取得30天資料
        data_30d = self.db_manager.get_recent_data(30)
        summary_30d = self.db_manager.get_daily_summary(30)
        
        if data_30d.empty:
            self.logger.warning("沒有30天內的資料可供分析")
            return None
        
        # 生成報告
        report = {
            "基本資訊": self.generate_basic_info(data_30d, summary_30d),
            "三大法人分析": self.generate_institutional_analysis(data_30d),
            "契約分析": self.generate_contract_analysis(data_30d),
            "趨勢分析": self.generate_trend_analysis(summary_30d),
            "警示提醒": self.generate_alerts(data_30d, summary_30d)
        }
        
        # 儲存JSON報告
        json_path = self.output_dir / f"台期所30日報告_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        
        # 生成Excel報告
        excel_path = self.output_dir / f"台期所30日報告_{timestamp}.xlsx"
        self.export_excel_report(data_30d, summary_30d, excel_path)
        
        # 生成圖表
        chart_dir = self.output_dir / f"charts_{timestamp}"
        self.generate_charts(data_30d, summary_30d, chart_dir)
        
        # 生成Markdown報告
        md_path = self.output_dir / f"台期所30日報告_{timestamp}.md"
        self.generate_markdown_report(report, md_path)
        
        self.logger.info(f"30天報告生成完成：{excel_path}")
        return report
    
    def generate_basic_info(self, data_30d, summary_30d):
        """生成基本資訊"""
        latest_date = data_30d['date'].max() if not data_30d.empty else "無資料"
        oldest_date = data_30d['date'].min() if not data_30d.empty else "無資料"
        
        return {
            "報告生成時間": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "資料期間": f"{oldest_date} ~ {latest_date}",
            "總交易日數": len(summary_30d) if not summary_30d.empty else 0,
            "涵蓋契約": list(data_30d['contract_code'].unique()) if not data_30d.empty else [],
            "資料筆數": len(data_30d),
            "最新更新": latest_date
        }
    
    def generate_institutional_analysis(self, data_30d):
        """生成三大法人分析"""
        if data_30d.empty:
            return {"錯誤": "無資料"}
        
        # 計算各法人平均淨部位
        institutional_summary = data_30d.groupby('identity_type')['net_position'].agg([
            'mean', 'sum', 'std', 'count'
        ]).round(2)
        
        # 最新一日的法人部位
        latest_date = data_30d['date'].max()
        latest_positions = data_30d[data_30d['date'] == latest_date].groupby('identity_type')['net_position'].sum()
        
        # 計算變動趨勢（與7天前比較）
        seven_days_ago = (datetime.strptime(latest_date, '%Y/%m/%d') - timedelta(days=7)).strftime('%Y/%m/%d')
        past_positions = data_30d[data_30d['date'] == seven_days_ago].groupby('identity_type')['net_position'].sum()
        
        trend_analysis = {}
        for identity in latest_positions.index:
            current = latest_positions.get(identity, 0)
            past = past_positions.get(identity, 0)
            change = current - past
            trend_analysis[identity] = {
                "最新部位": current,
                "7天前部位": past,
                "變動": change,
                "變動率": f"{(change/past*100 if past != 0 else 0):.2f}%"
            }
        
        return {
            "30天統計": institutional_summary.to_dict(),
            "最新部位": latest_positions.to_dict(),
            "7天變動分析": trend_analysis
        }
    
    def generate_contract_analysis(self, data_30d):
        """生成契約分析"""
        if data_30d.empty:
            return {"錯誤": "無資料"}
        
        # 各契約的活躍度分析
        contract_activity = data_30d.groupby('contract_code').agg({
            'long_position': 'sum',
            'short_position': 'sum',
            'net_position': ['sum', 'std'],
            'date': 'nunique'
        }).round(2)
        
        # 計算各契約的總交易量
        contract_activity['total_volume'] = contract_activity[('long_position', 'sum')] + contract_activity[('short_position', 'sum')]
        
        # 找出最活躍的契約
        most_active = contract_activity['total_volume'].idxmax() if not contract_activity.empty else "無"
        
        return {
            "契約活躍度統計": contract_activity.to_dict(),
            "最活躍契約": most_active,
            "契約數量": len(data_30d['contract_code'].unique())
        }
    
    def generate_trend_analysis(self, summary_30d):
        """生成趨勢分析"""
        if summary_30d.empty:
            return {"錯誤": "無資料"}
        
        # 計算移動平均
        summary_30d = summary_30d.sort_values('date')
        summary_30d['外資_MA7'] = summary_30d['foreign_net'].rolling(7).mean()
        summary_30d['自營商_MA7'] = summary_30d['dealer_net'].rolling(7).mean()
        summary_30d['投信_MA7'] = summary_30d['trust_net'].rolling(7).mean()
        
        # 趨勢判斷
        latest_trend = summary_30d.tail(1).iloc[0] if not summary_30d.empty else {}
        
        trends = {}
        for institution in ['foreign_net', 'dealer_net', 'trust_net']:
            recent_avg = summary_30d[institution].tail(7).mean()
            early_avg = summary_30d[institution].head(7).mean()
            
            if recent_avg > early_avg * 1.1:
                trend = "上升趨勢"
            elif recent_avg < early_avg * 0.9:
                trend = "下降趨勢"
            else:
                trend = "橫盤整理"
            
            trends[institution] = {
                "趨勢": trend,
                "近7日均值": round(recent_avg, 2),
                "前7日均值": round(early_avg, 2)
            }
        
        return {
            "趨勢判斷": trends,
            "最新資料": latest_trend.to_dict() if hasattr(latest_trend, 'to_dict') else latest_trend
        }
    
    def generate_alerts(self, data_30d, summary_30d):
        """生成警示提醒"""
        alerts = []
        
        if data_30d.empty or summary_30d.empty:
            return {"警示": ["無足夠資料進行警示分析"]}
        
        # 檢查是否有極端部位
        latest_summary = summary_30d.iloc[0] if not summary_30d.empty else {}
        
        # 外資極端部位警示
        if abs(latest_summary.get('foreign_net', 0)) > 50000:
            alerts.append(f"⚠️ 外資淨部位達 {latest_summary.get('foreign_net', 0):,} 口，屬極端部位")
        
        # 連續同向操作警示
        if len(summary_30d) >= 5:
            recent_foreign = summary_30d.head(5)['foreign_net'].tolist()
            if all(x > 0 for x in recent_foreign):
                alerts.append("🔺 外資連續5日淨多方，注意反轉風險")
            elif all(x < 0 for x in recent_foreign):
                alerts.append("🔻 外資連續5日淨空方，注意反彈機會")
        
        # 成交量異常警示
        avg_volume = summary_30d['total_volume'].mean()
        latest_volume = latest_summary.get('total_volume', 0)
        if latest_volume > avg_volume * 1.5:
            alerts.append("📈 最新交易量異常放大，注意市場變化")
        
        return {"警示": alerts if alerts else ["目前無特殊警示"]}
    
    def export_excel_report(self, data_30d, summary_30d, output_path):
        """匯出Excel報告"""
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 30天原始資料
            if not data_30d.empty:
                data_30d.to_excel(writer, sheet_name='30天原始資料', index=False)
            
            # 每日摘要
            if not summary_30d.empty:
                summary_30d.to_excel(writer, sheet_name='每日摘要', index=False)
            
            # 三大法人透視表
            if not data_30d.empty:
                pivot_table = pd.pivot_table(
                    data_30d, 
                    values='net_position',
                    index='date',
                    columns='identity_type',
                    aggfunc='sum'
                )
                pivot_table.to_excel(writer, sheet_name='三大法人透視表')
    
    def generate_charts(self, data_30d, summary_30d, chart_dir):
        """生成圖表"""
        chart_dir.mkdir(exist_ok=True)
        
        if summary_30d.empty:
            return
        
        # 三大法人淨部位趨勢圖
        plt.figure(figsize=(12, 8))
        plt.plot(summary_30d['date'], summary_30d['foreign_net'], label='外資', marker='o')
        plt.plot(summary_30d['date'], summary_30d['dealer_net'], label='自營商', marker='s')
        plt.plot(summary_30d['date'], summary_30d['trust_net'], label='投信', marker='^')
        
        plt.title('三大法人30天淨部位趨勢', fontsize=16)
        plt.xlabel('日期', fontsize=12)
        plt.ylabel('淨部位（口）', fontsize=12)
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(chart_dir / '三大法人趨勢.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 成交量變化圖
        plt.figure(figsize=(12, 6))
        plt.bar(summary_30d['date'], summary_30d['total_volume'], alpha=0.7)
        plt.title('30天成交量變化', fontsize=16)
        plt.xlabel('日期', fontsize=12)
        plt.ylabel('成交量', fontsize=12)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(chart_dir / '成交量變化.png', dpi=300, bbox_inches='tight')
        plt.close()
    
    def generate_markdown_report(self, report_data, output_path):
        """生成Markdown格式報告"""
        md_content = f"""# 台期所30天分析報告

## 📊 基本資訊
- **報告生成時間**: {report_data['基本資訊']['報告生成時間']}
- **資料期間**: {report_data['基本資訊']['資料期間']}
- **總交易日數**: {report_data['基本資訊']['總交易日數']} 日
- **資料筆數**: {report_data['基本資訊']['資料筆數']:,} 筆

## 🏛️ 三大法人分析

### 最新部位狀況
"""
        
        # 添加三大法人部位資訊
        if '最新部位' in report_data['三大法人分析']:
            for institution, position in report_data['三大法人分析']['最新部位'].items():
                md_content += f"- **{institution}**: {position:,} 口\n"
        
        md_content += f"""
## ⚠️ 重要警示

"""
        # 添加警示資訊
        for alert in report_data['警示提醒']['警示']:
            md_content += f"- {alert}\n"
        
        md_content += f"""
---
*本報告由台期所自動化分析系統生成*
"""
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)


if __name__ == "__main__":
    # 測試日報生成器
    generator = DailyReportGenerator()
    report = generator.generate_30day_report()
    print("30天日報生成完成！") 