#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
雲端同步設定腳本
自動將台期所資料同步到Google Drive/OneDrive
"""

import os
import shutil
from pathlib import Path
import json
import subprocess
import platform

class CloudSyncManager:
    """雲端同步管理器"""
    
    def __init__(self):
        self.system = platform.system()
        self.config_file = Path("config/cloud_sync.json")
        self.config = self.load_config()
    
    def load_config(self):
        """載入雲端同步設定"""
        if self.config_file.exists():
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # 建立預設設定
            default_config = {
                "google_drive": {
                    "enabled": False,
                    "folder_path": "台期所分析資料",
                    "sync_files": [
                        "output/台期所最新30天資料.xlsx",
                        "reports/台期所30日報告_*.xlsx",
                        "reports/台期所30日報告_*.md"
                    ]
                },
                "onedrive": {
                    "enabled": False,
                    "folder_path": "台期所分析資料",
                    "sync_files": [
                        "output/台期所最新30天資料.xlsx",
                        "reports/台期所30日報告_*.xlsx"
                    ]
                },
                "dropbox": {
                    "enabled": False,
                    "folder_path": "台期所分析資料"
                },
                "auto_sync": True,
                "sync_frequency": "daily"
            }
            
            self.config_file.parent.mkdir(exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            
            return default_config
    
    def detect_cloud_drives(self):
        """偵測系統中已安裝的雲端硬碟"""
        detected = {}
        
        if self.system == "Windows":
            # 偵測OneDrive
            onedrive_paths = [
                os.path.expanduser("~/OneDrive"),
                os.path.expanduser("~/OneDrive - Personal"),
                "C:/Users/{}/OneDrive".format(os.getenv('USERNAME', ''))
            ]
            
            for path in onedrive_paths:
                if os.path.exists(path):
                    detected['onedrive'] = path
                    break
            
            # 偵測Google Drive
            google_drive_paths = [
                os.path.expanduser("~/Google Drive"),
                "C:/Users/{}/Google Drive".format(os.getenv('USERNAME', ''))
            ]
            
            for path in google_drive_paths:
                if os.path.exists(path):
                    detected['google_drive'] = path
                    break
        
        elif self.system == "Darwin":  # macOS
            # 偵測iCloud Drive
            icloud_path = os.path.expanduser("~/Library/Mobile Documents/com~apple~CloudDocs")
            if os.path.exists(icloud_path):
                detected['icloud'] = icloud_path
            
            # 偵測Google Drive
            google_drive_path = os.path.expanduser("~/Google Drive")
            if os.path.exists(google_drive_path):
                detected['google_drive'] = google_drive_path
            
            # 偵測OneDrive
            onedrive_path = os.path.expanduser("~/OneDrive")
            if os.path.exists(onedrive_path):
                detected['onedrive'] = onedrive_path
        
        return detected
    
    def setup_onedrive_sync(self):
        """設定OneDrive同步"""
        detected = self.detect_cloud_drives()
        
        if 'onedrive' not in detected:
            print("❌ 未偵測到OneDrive，請先安裝OneDrive應用程式")
            return False
        
        onedrive_path = detected['onedrive']
        target_folder = os.path.join(onedrive_path, "台期所分析資料")
        
        # 建立目標資料夾
        os.makedirs(target_folder, exist_ok=True)
        
        # 複製關鍵檔案
        files_to_sync = [
            "output/台期所最新30天資料.xlsx",
            "data/taifex_data.db"
        ]
        
        for file_path in files_to_sync:
            if os.path.exists(file_path):
                target_path = os.path.join(target_folder, os.path.basename(file_path))
                shutil.copy2(file_path, target_path)
                print(f"✅ 已同步: {file_path} -> OneDrive")
        
        print(f"🎉 OneDrive同步設定完成！")
        print(f"📁 雲端資料夾: {target_folder}")
        print(f"🌐 現在可以從任何裝置存取資料了！")
        
        return True
    
    def setup_google_drive_sync(self):
        """設定Google Drive同步"""
        detected = self.detect_cloud_drives()
        
        if 'google_drive' not in detected:
            print("❌ 未偵測到Google Drive，請先安裝Google Drive應用程式")
            return False
        
        google_drive_path = detected['google_drive']
        target_folder = os.path.join(google_drive_path, "台期所分析資料")
        
        # 建立目標資料夾
        os.makedirs(target_folder, exist_ok=True)
        
        # 複製關鍵檔案
        files_to_sync = [
            "output/台期所最新30天資料.xlsx",
            "data/taifex_data.db"
        ]
        
        for file_path in files_to_sync:
            if os.path.exists(file_path):
                target_path = os.path.join(target_folder, os.path.basename(file_path))
                shutil.copy2(file_path, target_path)
                print(f"✅ 已同步: {file_path} -> Google Drive")
        
        print(f"🎉 Google Drive同步設定完成！")
        print(f"📁 雲端資料夾: {target_folder}")
        
        return True
    
    def auto_sync_files(self):
        """自動同步檔案到雲端"""
        detected = self.detect_cloud_drives()
        synced = False
        
        # 要同步的檔案
        files_to_sync = [
            "output/台期所最新30天資料.xlsx",
            "data/taifex_data.db"
        ]
        
        # 尋找最新的報告檔案
        reports_dir = Path("reports")
        if reports_dir.exists():
            latest_report = None
            for report_file in reports_dir.glob("台期所30日報告_*.xlsx"):
                if latest_report is None or report_file.stat().st_mtime > latest_report.stat().st_mtime:
                    latest_report = report_file
            
            if latest_report:
                files_to_sync.append(str(latest_report))
        
        # 同步到OneDrive
        if 'onedrive' in detected and self.config.get('onedrive', {}).get('enabled', False):
            target_folder = os.path.join(detected['onedrive'], "台期所分析資料")
            os.makedirs(target_folder, exist_ok=True)
            
            for file_path in files_to_sync:
                if os.path.exists(file_path):
                    target_path = os.path.join(target_folder, os.path.basename(file_path))
                    shutil.copy2(file_path, target_path)
                    synced = True
        
        # 同步到Google Drive
        if 'google_drive' in detected and self.config.get('google_drive', {}).get('enabled', False):
            target_folder = os.path.join(detected['google_drive'], "台期所分析資料")
            os.makedirs(target_folder, exist_ok=True)
            
            for file_path in files_to_sync:
                if os.path.exists(file_path):
                    target_path = os.path.join(target_folder, os.path.basename(file_path))
                    shutil.copy2(file_path, target_path)
                    synced = True
        
        return synced
    
    def create_web_dashboard(self):
        """建立簡單的網頁版儀表板"""
        html_content = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>台期所資料分析儀表板</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; }
        .header { text-align: center; color: #333; border-bottom: 2px solid #007acc; padding-bottom: 10px; }
        .section { margin: 20px 0; padding: 15px; border-left: 4px solid #007acc; background: #f9f9f9; }
        .download-link { display: inline-block; margin: 10px; padding: 10px 20px; background: #007acc; color: white; text-decoration: none; border-radius: 5px; }
        .download-link:hover { background: #005580; }
        .status { padding: 5px 10px; border-radius: 15px; font-size: 12px; }
        .online { background: #4CAF50; color: white; }
        .offline { background: #f44336; color: white; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏛️ 台期所資料分析儀表板</h1>
            <p>最後更新: <span id="lastUpdate"></span> <span class="status offline">電腦離線</span></p>
        </div>
        
        <div class="section">
            <h3>📊 快速存取</h3>
            <a href="台期所最新30天資料.xlsx" class="download-link">📄 最新30天資料</a>
            <a href="../reports/" class="download-link">📈 分析報告</a>
            <a href="https://github.com/你的用戶名/你的倉庫名" class="download-link">🔗 GitHub倉庫</a>
        </div>
        
        <div class="section">
            <h3>📱 行動裝置使用說明</h3>
            <ol>
                <li>將此網頁加入書籤</li>
                <li>透過雲端硬碟APP開啟Excel檔案</li>
                <li>或直接訪問GitHub網頁版檢視資料</li>
            </ol>
        </div>
        
        <div class="section">
            <h3>🔗 重要連結</h3>
            <ul>
                <li><a href="https://github.com/你的用戶名/你的倉庫名/tree/main/output">GitHub - 輸出檔案</a></li>
                <li><a href="https://github.com/你的用戶名/你的倉庫名/tree/main/reports">GitHub - 分析報告</a></li>
                <li><a href="https://github.com/你的用戶名/你的倉庫名/actions">GitHub Actions - 執行記錄</a></li>
            </ul>
        </div>
    </div>
    
    <script>
        document.getElementById('lastUpdate').textContent = new Date().toLocaleString('zh-TW');
    </script>
</body>
</html>
        '''
        
        # 儲存到雲端硬碟資料夾
        detected = self.detect_cloud_drives()
        saved = False
        
        for drive_name, drive_path in detected.items():
            target_folder = os.path.join(drive_path, "台期所分析資料")
            os.makedirs(target_folder, exist_ok=True)
            
            html_path = os.path.join(target_folder, "台期所儀表板.html")
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            print(f"✅ 網頁儀表板已建立: {html_path}")
            saved = True
        
        return saved


def main():
    """主程序"""
    print("🌐 台期所雲端同步設定工具")
    print("=" * 50)
    
    sync_manager = CloudSyncManager()
    
    # 偵測雲端硬碟
    detected = sync_manager.detect_cloud_drives()
    
    if not detected:
        print("❌ 未偵測到任何雲端硬碟服務")
        print("\n建議安裝以下任一服務：")
        print("- OneDrive (Windows內建)")
        print("- Google Drive")
        print("- iCloud Drive (Mac)")
        return
    
    print("✅ 偵測到以下雲端服務：")
    for drive_name, drive_path in detected.items():
        print(f"  📁 {drive_name.title()}: {drive_path}")
    
    print("\n選擇操作：")
    print("1. 設定OneDrive同步")
    print("2. 設定Google Drive同步")
    print("3. 執行自動同步")
    print("4. 建立網頁儀表板")
    print("5. 全部設定")
    
    choice = input("\n請選擇 (1-5): ").strip()
    
    if choice == '1' and 'onedrive' in detected:
        sync_manager.setup_onedrive_sync()
    elif choice == '2' and 'google_drive' in detected:
        sync_manager.setup_google_drive_sync()
    elif choice == '3':
        if sync_manager.auto_sync_files():
            print("✅ 檔案同步完成")
        else:
            print("❌ 同步失敗，請先設定雲端服務")
    elif choice == '4':
        if sync_manager.create_web_dashboard():
            print("✅ 網頁儀表板建立完成")
    elif choice == '5':
        # 全部設定
        if 'onedrive' in detected:
            sync_manager.setup_onedrive_sync()
        if 'google_drive' in detected:
            sync_manager.setup_google_drive_sync()
        sync_manager.create_web_dashboard()
        print("\n🎉 所有設定完成！現在可以從任何裝置存取台期所資料了！")
    else:
        print("❌ 無效選擇或缺少對應的雲端服務")


if __name__ == "__main__":
    main() 