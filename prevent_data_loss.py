#!/usr/bin/env python3
"""
資料保護監控工具
定期檢查並防止意外的資料丟失
"""

import sqlite3
import pandas as pd
from pathlib import Path
import json
import shutil
from datetime import datetime
from google_sheets_manager import GoogleSheetsManager

class DataProtectionMonitor:
    """資料保護監控器"""
    
    def __init__(self):
        self.backup_dir = Path("backup")
        self.backup_dir.mkdir(exist_ok=True)
        
    def create_daily_backup(self):
        """創建每日資料庫備份"""
        try:
            db_path = Path("data/taifex_data.db")
            if db_path.exists():
                backup_name = f"taifex_data_backup_{datetime.now().strftime('%Y%m%d')}.db"
                backup_path = self.backup_dir / backup_name
                
                shutil.copy2(db_path, backup_path)
                print(f"✅ 每日備份完成: {backup_name}")
                
                # 清理7天前的備份
                self.cleanup_old_backups()
                return True
            else:
                print("❌ 找不到資料庫檔案")
                return False
                
        except Exception as e:
            print(f"❌ 備份失敗: {e}")
            return False
    
    def cleanup_old_backups(self, keep_days=7):
        """清理舊備份檔案"""
        try:
            backup_files = list(self.backup_dir.glob("taifex_data_backup_*.db"))
            backup_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            
            # 保留最近N天的備份
            if len(backup_files) > keep_days:
                for old_backup in backup_files[keep_days:]:
                    old_backup.unlink()
                    print(f"🗑️ 清理舊備份: {old_backup.name}")
                    
        except Exception as e:
            print(f"⚠️ 清理舊備份失敗: {e}")
    
    def check_sheets_data_integrity(self):
        """檢查Google Sheets資料完整性"""
        try:
            config_file = Path("config/spreadsheet_config.json")
            if not config_file.exists():
                print("❌ 找不到Google Sheets設定檔")
                return False
            
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            sheets_manager = GoogleSheetsManager()
            sheets_manager.connect_spreadsheet(config['spreadsheet_id'])
            
            if sheets_manager.spreadsheet:
                worksheet = sheets_manager.spreadsheet.worksheet("歷史資料")
                all_data = worksheet.get_all_values()
                current_rows = len(all_data)
                
                print(f"📊 Google Sheets目前狀況:")
                print(f"  - 總行數: {current_rows}")
                
                if current_rows < 100:
                    print("⚠️ 警告: Google Sheets資料量異常少，可能需要檢查")
                    return False
                elif current_rows > 9500:
                    print("⚠️ 警告: Google Sheets接近行數限制，建議整理資料")
                    return False
                else:
                    print("✅ Google Sheets資料量正常")
                    return True
            else:
                print("❌ 無法連接到Google Sheets")
                return False
                
        except Exception as e:
            print(f"❌ 檢查Google Sheets失敗: {e}")
            return False
    
    def check_database_integrity(self):
        """檢查資料庫完整性"""
        try:
            db_path = Path("data/taifex_data.db")
            if not db_path.exists():
                print("❌ 找不到資料庫檔案")
                return False
            
            conn = sqlite3.connect(db_path)
            
            # 檢查總筆數
            query = "SELECT COUNT(*) as total FROM futures_data"
            result = pd.read_sql_query(query, conn)
            total_records = result.iloc[0]['total']
            
            # 檢查最近的資料
            recent_query = """
                SELECT MAX(date) as latest_date, COUNT(*) as recent_count
                FROM futures_data 
                WHERE date >= date('now', '-7 days')
            """
            recent_result = pd.read_sql_query(recent_query, conn)
            
            conn.close()
            
            print(f"📊 資料庫狀況:")
            print(f"  - 總筆數: {total_records:,}")
            print(f"  - 最新日期: {recent_result.iloc[0]['latest_date']}")
            print(f"  - 近7天資料: {recent_result.iloc[0]['recent_count']} 筆")
            
            if total_records == 0:
                print("❌ 警告: 資料庫為空")
                return False
            elif recent_result.iloc[0]['recent_count'] == 0:
                print("⚠️ 警告: 近期沒有新資料")
                return False
            else:
                print("✅ 資料庫狀況正常")
                return True
                
        except Exception as e:
            print(f"❌ 檢查資料庫失敗: {e}")
            return False
    
    def generate_protection_report(self):
        """生成保護狀況報告"""
        print("🛡️ 資料保護狀況檢查")
        print("=" * 50)
        
        # 檢查各項狀況
        backup_ok = self.create_daily_backup()
        db_ok = self.check_database_integrity()
        sheets_ok = self.check_sheets_data_integrity()
        
        # 生成報告
        print(f"\n📋 保護狀況摘要:")
        print(f"  - 備份機制: {'✅ 正常' if backup_ok else '❌ 異常'}")
        print(f"  - 資料庫: {'✅ 正常' if db_ok else '❌ 異常'}")
        print(f"  - Google Sheets: {'✅ 正常' if sheets_ok else '❌ 異常'}")
        
        if backup_ok and db_ok and sheets_ok:
            print("\n🎉 所有資料保護機制運作正常！")
        else:
            print("\n⚠️ 發現問題，建議立即檢查並修復")
            
        return backup_ok and db_ok and sheets_ok

def main():
    """主程式"""
    monitor = DataProtectionMonitor()
    monitor.generate_protection_report()

if __name__ == "__main__":
    main() 