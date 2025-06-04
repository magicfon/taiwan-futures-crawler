#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
獲取Google Sheets的網址
"""

from google_sheets_manager import GoogleSheetsManager

def get_sheet_url():
    """獲取Google Sheets的網址"""
    try:
        print('🔗 獲取Google Sheets網址...')
        
        gm = GoogleSheetsManager()
        spreadsheet = gm.client.open('台期所資料分析')
        
        # 獲取試算表資訊
        sheet_id = spreadsheet.id
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
        
        print(f'📊 試算表名稱: {spreadsheet.title}')
        print(f'🆔 試算表ID: {sheet_id}')
        print(f'🔗 試算表網址: {sheet_url}')
        
        # 獲取所有工作表
        worksheets = spreadsheet.worksheets()
        print(f'\n📋 包含的工作表:')
        for i, ws in enumerate(worksheets, 1):
            ws_url = f"{sheet_url}#gid={ws.id}"
            print(f'  {i}. {ws.title} - {ws_url}')
        
        return sheet_url
        
    except Exception as e:
        print(f'❌ 獲取網址失敗: {e}')
        return None

if __name__ == "__main__":
    get_sheet_url() 