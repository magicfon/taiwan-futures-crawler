#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
台灣期貨交易所資料爬取工具 (Python 版本)
可用於快速爬取整年的歷史資料
"""

import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
import time
import argparse
import datetime
import re
import concurrent.futures
from tqdm import tqdm
import logging
import random
import json
from pathlib import Path
try:
    from database_manager import TaifexDatabaseManager
    from daily_report_generator import DailyReportGenerator
    DB_AVAILABLE = True
except ImportError:
    DB_AVAILABLE = False
    print("警告: 資料庫模組未找到，將使用傳統檔案存儲方式")

# 設定基本參數
BASE_URL = 'https://www.taifex.com.tw/cht/3/futContractsDate'
# 可爬取的期貨契約代碼: 台指期, 電子期, 小型台指期, 微型台指期, 美國那斯達克100期貨
CONTRACTS = ['TX', 'TE', 'MTX', 'ZMX', 'NQF']
# 契約名稱對照表
CONTRACT_NAMES = {
    'TX': '臺股期貨',
    'TE': '電子期貨',
    'MTX': '小型臺指期貨',
    'ZMX': '微型臺指期貨',
    'NQF': '美國那斯達克100期貨'
}
# 身份別列表
IDENTITIES = ['自營商', '投信', '外資']

# 設定日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("taifex_crawler.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("期貨爬蟲")

class TaifexCrawler:
    def __init__(self, output_dir="output", max_retries=3, delay=0.5, 
                 max_workers=10, timeout=30, use_proxy=False):
        """
        初始化爬蟲
        
        Args:
            output_dir: 輸出目錄
            max_retries: 最大重試次數
            delay: 請求間隔時間 (秒)
            max_workers: 最大工作線程數
            timeout: 請求超時時間 (秒)
            use_proxy: 是否使用代理
        """
        self.output_dir = output_dir
        self.max_retries = max_retries
        self.delay = delay
        self.max_workers = max_workers
        self.timeout = timeout
        self.use_proxy = use_proxy
        self.session = requests.Session()
        
        # 設定請求標頭
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
            'Connection': 'keep-alive',
            'Referer': 'https://www.taifex.com.tw/cht/3/futDailyMarketView',
        }
        
        # 確保輸出目錄存在
        os.makedirs(self.output_dir, exist_ok=True)
    
    def fetch_data(self, date_str, contract, query_type='2', market_code='0', identity=None):
        """
        從期交所獲取特定日期和契約的資料
        
        Args:
            date_str: 日期字串，格式為 'yyyy/MM/dd'
            contract: 契約代碼 (TX, TE, MTX, ZMX, NQF)
            query_type: 查詢類型 ('1', '2', '3')
            market_code: 市場代碼 ('0', '1', '2')
            identity: 身份別 (自營商, 投信, 外資)
            
        Returns:
            解析後的資料字典或 None (如果無資料)
        """
        # 檢查日期是否超過今天
        try:
            target_date = datetime.datetime.strptime(date_str, "%Y/%m/%d")
            today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            
            if target_date > today:
                logger.warning(f"日期 {date_str} 超過今天，跳過爬取")
                return None
        except Exception as e:
            logger.error(f"日期格式檢查錯誤: {str(e)}")
            # 忽略錯誤，繼續執行
        
        for retry in range(self.max_retries):
            try:
                # 構建查詢參數
                params = {
                    'queryType': query_type,
                    'marketCode': market_code,
                    'dateaddcnt': '',
                    'commodity_id': contract,
                    'queryDate': date_str
                }
                
                # 添加隨機延遲，避免被封IP
                time.sleep(self.delay + random.uniform(0, 0.5))
                
                # 發送 GET 請求
                response = self.session.get(
                    BASE_URL, 
                    params=params, 
                    headers=self.headers,
                    timeout=self.timeout
                )
                
                # 檢查響應狀態
                if response.status_code != 200:
                    logger.warning(f"請求失敗: {response.status_code}, 重試中 ({retry+1}/{self.max_retries})")
                    continue
                
                # 檢查是否有錯誤訊息或無資料訊息
                html_content = response.text
                if self._has_no_data_message(html_content):
                    logger.info(f"{date_str} {contract} 無交易資料")
                    return None
                
                if self._has_error_message(html_content):
                    logger.warning(f"{date_str} {contract} 頁面包含錯誤訊息")
                    continue
                
                # 解析資料
                if identity:
                    data = self._parse_identity_data(html_content, contract, date_str, identity)
                    if not data:
                        logger.warning(f"{date_str} {contract} 找不到身份別 {identity} 的資料")
                else:
                    data = self._parse_contract_data(html_content, contract, date_str)
                    if not data:
                        logger.warning(f"{date_str} {contract} 找不到契約資料")
                
                if data:
                    return data
                
                logger.warning(f"{date_str} {contract} 資料解析失敗，重試中 ({retry+1}/{self.max_retries})")
            
            except Exception as e:
                logger.error(f"爬取 {date_str} {contract} 資料時發生錯誤: {str(e)}")
                if retry < self.max_retries - 1:
                    logger.info(f"等待後重試 ({retry+1}/{self.max_retries})")
                    time.sleep(self.delay * 2)
        
        return None
    
    def _has_no_data_message(self, html_content):
        """檢查網頁是否顯示「無交易資料」訊息"""
        no_data_patterns = [
            '查無資料', '無交易資料', '尚無資料', '無此資料',
            'No data', '不存在', '不提供', '維護中'
        ]
        return any(pattern in html_content for pattern in no_data_patterns)
    
    def _has_error_message(self, html_content):
        """檢查網頁是否顯示錯誤訊息"""
        error_patterns = [
            '系統發生錯誤', '查詢發生錯誤', '日期錯誤', 
            '請輸入正確的日期', 'Error occurred', '資料錯誤',
            '錯誤代碼', '伺服器錯誤', 'Server Error', '無法處理您的請求'
        ]
        return any(pattern in html_content for pattern in error_patterns)
    
    def _parse_number(self, text):
        """解析表格中的數字，處理千位分隔符和小數點"""
        if not text:
            return 0
        
        try:
            # 移除所有非數字字符（保留減號表示負數和小數點）
            cleaned_str = re.sub(r'[^\d\-\.]', '', str(text))
            
            # 檢查是否有小數點
            if '.' in cleaned_str:
                # 解析為浮點數
                return float(cleaned_str) if cleaned_str else 0
            else:
                # 轉換為整數
                return int(cleaned_str) if cleaned_str else 0
        except:
            # 如果解析失敗，嘗試額外的處理（例如科學記數法）
            try:
                return float(text.replace(',', ''))
            except:
                return 0
    
    def _parse_contract_data(self, html_content, contract, date_str):
        """解析基本契約資料"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 尋找主要資料表格
            tables = soup.find_all('table', {'class': 'table_f'})
            if not tables:
                tables = soup.find_all('table')
            
            if not tables:
                return None
            
            # 選擇最大的表格
            main_table = max(tables, key=lambda t: len(str(t)))
            
            # 尋找契約所在行
            target_row = None
            highest_score = 0
            
            # 契約模式匹配字典
            contract_patterns = {
                'TX': ['臺指期', '台指期', '臺股期', '台股期'],
                'TE': ['電子期'],
                'MTX': ['小型臺指', '小型台指', '小臺指', '小台指'],
                'ZMX': ['微型臺指', '微型台指', '微臺指', '微台指'],
                'NQF': ['那斯達克', 'Nasdaq', '美國那斯達克'],
            }
            
            for row in main_table.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                if not cells:
                    continue
                
                row_text = ' '.join([cell.get_text(strip=True) for cell in cells])
                score = 0
                
                # 精確匹配契約代碼
                if contract in row_text:
                    score += 10
                if CONTRACT_NAMES.get(contract, '') in row_text:
                    score += 10
                
                # 針對特定契約的模糊匹配
                if contract in contract_patterns:
                    for pattern in contract_patterns[contract]:
                        if pattern in row_text:
                            score += 8
                            break
                
                # 避免交叉匹配到錯誤的契約
                for c, patterns in contract_patterns.items():
                    if c != contract:  # 不是目標契約
                        for pattern in patterns:
                            if pattern in row_text:
                                score -= 10  # 懲罰非目標契約匹配
                                break
                
                # 檢查是否包含數字（交易數據）
                if re.search(r'\d', row_text):
                    score += 2
                
                # 記錄高分匹配，幫助調試
                if score >= 10:
                    logger.debug(f"契約 {contract} 匹配行分數: {score}, 行文本: {row_text[:30]}...")
                
                if score > highest_score:
                    highest_score = score
                    target_row = row
            
            if not target_row or highest_score < 10:
                return None
            
            # 解析行中的資料
            cells = target_row.find_all('td')
            if len(cells) < 6:
                return None
            
            # 獲取所有單元格文本
            cell_texts = [cell.get_text(strip=True) for cell in cells]
            
            # 檢測表格格式
            has_identity = False
            identity = ''
            
            # 檢查是否為新格式表格（包含身份別）
            for i, text in enumerate(cell_texts):
                if i < 3 and text in IDENTITIES:
                    has_identity = True
                    identity = text
                    break
            
            # 嘗試使用表頭查找欄位位置
            column_indices = self._find_column_indices(main_table)
            
            # 指定預設欄位位置
            indices = {
                '多方交易口數': 3,
                '多方契約金額': 4,
                '空方交易口數': 5,
                '空方契約金額': 6,
                '多空淨額交易口數': 7,
                '多空淨額契約金額': 8,
                '多方未平倉口數': 9,
                '多方未平倉契約金額': 10,
                '空方未平倉口數': 11,
                '空方未平倉契約金額': 12,
                '多空淨額未平倉口數': 13,
                '多空淨額未平倉契約金額': 14
            }
            
            # 如果找到表頭定義的欄位，則使用
            if column_indices:
                indices.update(column_indices)
            else:
                # 根據表格結構調整預設欄位位置
                if len(cell_texts) < 14:
                    # 精簡格式表格 (只有口數資料)
                    indices = {
                        '多方交易口數': 2,
                        '多方契約金額': -1,
                        '空方交易口數': 3,
                        '空方契約金額': -1,
                        '多空淨額交易口數': 4,
                        '多空淨額契約金額': -1,
                        '多方未平倉口數': 5,
                        '多方未平倉契約金額': -1,
                        '空方未平倉口數': 6,
                        '空方未平倉契約金額': -1,
                        '多空淨額未平倉口數': 7,
                        '多空淨額未平倉契約金額': -1
                    }
            
            # 安全獲取數值
            def safe_get(field):
                idx = indices.get(field, -1)
                if 0 <= idx < len(cell_texts):
                    return self._parse_number(cell_texts[idx])
                return 0
            
            # 構建資料字典
            data = {
                '日期': date_str,
                '契約名稱': contract,
                '身份別': identity,
                '多方交易口數': safe_get('多方交易口數'),
                '多方契約金額': safe_get('多方契約金額'),
                '空方交易口數': safe_get('空方交易口數'),
                '空方契約金額': safe_get('空方契約金額'),
                '多空淨額交易口數': safe_get('多空淨額交易口數'),
                '多空淨額契約金額': safe_get('多空淨額契約金額'),
                '多方未平倉口數': safe_get('多方未平倉口數'),
                '多方未平倉契約金額': safe_get('多方未平倉契約金額'),
                '空方未平倉口數': safe_get('空方未平倉口數'),
                '空方未平倉契約金額': safe_get('空方未平倉契約金額'),
                '多空淨額未平倉口數': safe_get('多空淨額未平倉口數'),
                '多空淨額未平倉契約金額': safe_get('多空淨額未平倉契約金額')
            }
            
            return data
            
        except Exception as e:
            logger.error(f"解析 {date_str} {contract} 資料時發生錯誤: {str(e)}")
            return None
    
    def _parse_identity_data(self, html_content, contract, date_str, identity):
        """解析特定身份別的資料"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 尋找主要資料表格
            tables = soup.find_all('table', {'class': 'table_f'})
            if not tables:
                tables = soup.find_all('table')
            
            if not tables:
                return None
            
            # 選擇最大的表格
            main_table = max(tables, key=lambda t: len(str(t)))
            
            # 先找到對應契約的自營商行
            base_row = None
            rows = main_table.find_all('tr')
            
            # 針對不同的契約使用不同的關鍵詞
            contract_keywords = {
                'TX': ['臺指期', '台指期', '臺股期', '台股期'],
                'TE': ['電子期'],
                'MTX': ['小型臺指', '小型台指', '小臺指', '小台指'],
                'ZMX': ['微型臺指', '微型台指', '微臺指', '微台指'],
                'NQF': ['那斯達克', 'Nasdaq', '美國那斯達克'],
            }
            
            # 找到自營商行
            for i, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                if not cells:
                    continue
                
                row_text = ' '.join([cell.get_text(strip=True) for cell in cells])
                
                # 精確匹配特定契約
                if contract in row_text or CONTRACT_NAMES.get(contract, '') in row_text:
                    # 確認是否包含自營商身份別
                    if '自營商' in row_text and (
                        contract in row_text or 
                        any(keyword in row_text for keyword in contract_keywords.get(contract, []))
                    ):
                        base_row = i
                        break
            
            if base_row is None:
                logger.warning(f"無法找到 {contract} 自營商行，無法確定其他身份別位置")
                return None
            
            # 根據身份別確定要取的行
            target_index = None
            if identity == '自營商':
                target_index = base_row
            elif identity == '投信':
                target_index = base_row + 1
            elif identity == '外資':
                target_index = base_row + 2
            else:
                return None
            
            # 確保目標索引有效
            if target_index < 0 or target_index >= len(rows):
                logger.warning(f"{contract} {identity} 目標行索引超出範圍: {target_index}")
                return None
            
            # 取得目標行
            target_row = rows[target_index]
            
            # 解析行中的資料
            cells = target_row.find_all('td')
            if len(cells) < 5:
                return None
            
            # 獲取所有單元格文本
            cell_texts = [cell.get_text(strip=True) for cell in cells]
            
            # 嘗試使用表頭查找欄位位置
            column_indices = self._find_column_indices(main_table)
            
            # 檢查是否為第一欄包含身份別的格式
            is_identity_first = cell_texts[0] == identity or identity in cell_texts[0]
            
            # 指定預設欄位位置
            if is_identity_first:
                # 如果身份別在第一欄
                indices = {
                    '多方交易口數': 1,
                    '多方契約金額': 2,
                    '空方交易口數': 3,
                    '空方契約金額': 4,
                    '多空淨額交易口數': 5,
                    '多空淨額契約金額': 6,
                    '多方未平倉口數': 7,
                    '多方未平倉契約金額': 8,
                    '空方未平倉口數': 9,
                    '空方未平倉契約金額': 10,
                    '多空淨額未平倉口數': 11,
                    '多空淨額未平倉契約金額': 12
                }
            else:
                # 標準格式
                indices = {
                    '多方交易口數': 3,
                    '多方契約金額': 4,
                    '空方交易口數': 5,
                    '空方契約金額': 6,
                    '多空淨額交易口數': 7,
                    '多空淨額契約金額': 8,
                    '多方未平倉口數': 9,
                    '多方未平倉契約金額': 10,
                    '空方未平倉口數': 11,
                    '空方未平倉契約金額': 12,
                    '多空淨額未平倉口數': 13,
                    '多空淨額未平倉契約金額': 14
                }
                
                # 如果列數不足，嘗試更精簡的格式
                if len(cell_texts) < 14:
                    indices = {
                        '多方交易口數': 2,
                        '多方契約金額': -1,
                        '空方交易口數': 3,
                        '空方契約金額': -1,
                        '多空淨額交易口數': 4,
                        '多空淨額契約金額': -1,
                        '多方未平倉口數': 5,
                        '多方未平倉契約金額': -1,
                        '空方未平倉口數': 6,
                        '空方未平倉契約金額': -1,
                        '多空淨額未平倉口數': 7,
                        '多空淨額未平倉契約金額': -1
                    }
            
            # 如果找到表頭定義的欄位，則使用
            if column_indices:
                indices.update(column_indices)
            
            # 安全獲取數值
            def safe_get(field):
                idx = indices.get(field, -1)
                if 0 <= idx < len(cell_texts):
                    return self._parse_number(cell_texts[idx])
                return 0
            
            # 構建資料字典
            data = {
                '日期': date_str,
                '契約名稱': contract,
                '身份別': identity,
                '多方交易口數': safe_get('多方交易口數'),
                '多方契約金額': safe_get('多方契約金額'),
                '空方交易口數': safe_get('空方交易口數'),
                '空方契約金額': safe_get('空方契約金額'),
                '多空淨額交易口數': safe_get('多空淨額交易口數'),
                '多空淨額契約金額': safe_get('多空淨額契約金額'),
                '多方未平倉口數': safe_get('多方未平倉口數'),
                '多方未平倉契約金額': safe_get('多方未平倉契約金額'),
                '空方未平倉口數': safe_get('空方未平倉口數'),
                '空方未平倉契約金額': safe_get('空方未平倉契約金額'),
                '多空淨額未平倉口數': safe_get('多空淨額未平倉口數'),
                '多空淨額未平倉契約金額': safe_get('多空淨額未平倉契約金額')
            }
            
            # 記錄日誌信息
            logger.debug(f"使用絕對位置解析 {contract} {identity}，在第 {target_index} 行")
            
            return data
            
        except Exception as e:
            logger.error(f"解析 {date_str} {contract} {identity} 資料時發生錯誤: {str(e)}")
            return None

    def crawl_single_day(self, date_str, contracts=None, identities=None):
        """爬取單日的資料"""
        if not contracts:
            contracts = CONTRACTS
        
        results = []
        
        # 爬取基本資料（非身份別）
        if not identities:
            for contract in contracts:
                data = self.fetch_data(date_str, contract)
                if data:
                    results.append(data)
        
        # 爬取多身份別資料
        else:
            for contract in contracts:
                for identity in identities:
                    data = self.fetch_data(date_str, contract, identity=identity)
                    if data:
                        results.append(data)
        
        return results

    def _is_business_day(self, date):
        """檢查是否為交易日（非週末）"""
        day = date.weekday()
        # 0-4 是週一到週五，5-6 是週末
        return day < 5

    def crawl_date_range(self, start_date, end_date, contracts=None, identities=None):
        """
        爬取指定日期範圍的資料
        
        Args:
            start_date: 開始日期 (datetime object)
            end_date: 結束日期 (datetime object)
            contracts: 要爬取的契約列表，默認為所有契約
            identities: 要爬取的身份別列表，默認為不爬取身份別資料
            
        Returns:
            DataFrame 包含所有爬取的資料
        """
        if not contracts:
            contracts = CONTRACTS
        
        # 確保結束日期不超過今天
        today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if end_date > today:
            logger.warning(f"結束日期 {end_date.strftime('%Y/%m/%d')} 超過今天，將設定為今天 {today.strftime('%Y/%m/%d')}")
            end_date = today
        
        # 如果開始日期超過今天，則無法爬取
        if start_date > today:
            logger.warning(f"開始日期 {start_date.strftime('%Y/%m/%d')} 超過今天，無法爬取資料")
            return pd.DataFrame()
        
        # 計算總任務數量（用於進度條）
        date_range = [start_date + datetime.timedelta(days=x) for x in range((end_date - start_date).days + 1)]
        # 過濾只包含今天之前的日期
        date_range = [d for d in date_range if d <= today]
        business_days = [d for d in date_range if self._is_business_day(d)]
        
        total_tasks = len(business_days) * len(contracts)
        if identities:
            total_tasks = len(business_days) * len(contracts) * len(identities)
        
        logger.info(f"開始爬取從 {start_date.strftime('%Y/%m/%d')} 到 {end_date.strftime('%Y/%m/%d')} 的資料")
        logger.info(f"共 {len(business_days)} 個交易日，{len(contracts)} 個契約類型")
        if identities:
            logger.info(f"包含 {len(identities)} 種身份別資料")
        
        # 準備任務列表
        tasks = []
        for date in business_days:
            date_str = date.strftime('%Y/%m/%d')
            
            if identities:
                for contract in contracts:
                    for identity in identities:
                        tasks.append((date_str, contract, identity))
            else:
                for contract in contracts:
                    tasks.append((date_str, contract, None))
        
        # 使用多線程加速爬取
        all_results = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # 創建任務
            future_to_task = {}
            for task in tasks:
                date_str, contract, identity = task
                future = executor.submit(self.fetch_data, date_str, contract, identity=identity)
                future_to_task[future] = task
            
            # 使用 tqdm 顯示進度條
            for future in tqdm(concurrent.futures.as_completed(future_to_task), 
                              total=len(future_to_task), 
                              desc="爬取進度"):
                date_str, contract, identity = future_to_task[future]
                try:
                    data = future.result()
                    if data:
                        all_results.append(data)
                        status = "成功"
                    else:
                        status = "無資料"
                except Exception as e:
                    logger.error(f"處理任務 {date_str} {contract} {identity} 時發生錯誤: {str(e)}")
                    status = f"錯誤: {str(e)}"
                
                logger.debug(f"{date_str} {contract} {identity or '總計'}: {status}")
        
        # 將結果轉換為 DataFrame
        if all_results:
            df = pd.DataFrame(all_results)
            return df
        else:
            logger.warning("沒有找到任何資料")
            return pd.DataFrame()

    def save_data(self, df, filename=None):
        """保存數據到 CSV 和 Excel 文件"""
        if df.empty:
            logger.warning("沒有資料可保存")
            return
        
        # 進行資料一致性檢查
        self._verify_data_consistency(df)
        
        # 如果未指定文件名，使用當前時間生成
        if not filename:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"taifex_data_{timestamp}"
        
        # 保存為 CSV
        csv_path = os.path.join(self.output_dir, f"{filename}.csv")
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        # 保存為 Excel
        excel_path = os.path.join(self.output_dir, f"{filename}.xlsx")
        df.to_excel(excel_path, index=False)
        
        logger.info(f"資料已保存至 {csv_path} 和 {excel_path}")
        return csv_path, excel_path
        
    def _verify_data_consistency(self, df):
        """驗證資料一致性，檢查是否有契約類型混淆的情況"""
        if '契約名稱' not in df.columns or '身份別' not in df.columns:
            return
            
        # 檢查每個身份別和契約的組合
        for identity in df['身份別'].dropna().unique():
            if not identity:
                continue
                
            for contract in CONTRACTS:
                contract_data = df[(df['身份別'] == identity) & (df['契約名稱'] == contract)]
                
                if len(contract_data) == 0:
                    continue
                    
                # 檢查數據是否合理
                has_suspicious_data = False
                
                # 如果同一契約有多筆資料，可能是錯誤識別
                if len(contract_data) > len(contract_data['日期'].unique()):
                    logger.warning(f"偵測到 {contract} {identity} 可能有重複資料")
                    has_suspicious_data = True
                
                # 檢查口數差距是否過大（可能是不同契約的資料被錯誤識別）
                if len(contract_data) > 3:
                    avg_vol = contract_data['多方交易口數'].mean()
                    max_vol = contract_data['多方交易口數'].max()
                    
                    if max_vol > 5 * avg_vol and max_vol > 1000:
                        logger.warning(f"偵測到 {contract} {identity} 可能有異常值，最大值 {max_vol} 遠大於平均值 {avg_vol}")
                        has_suspicious_data = True
                
                if has_suspicious_data:
                    logger.info(f"請檢查 {contract} {identity} 的資料是否正確，可能存在契約識別錯誤")

    def _find_column_indices(self, table):
        """尋找表頭中各欄位的索引位置"""
        # 欄位標題對應表
        column_keys = {
            '多方交易': {'vol': '多方交易口數', 'val': '多方契約金額'},
            '買方交易': {'vol': '多方交易口數', 'val': '多方契約金額'},
            '買進交易': {'vol': '多方交易口數', 'val': '多方契約金額'},
            '空方交易': {'vol': '空方交易口數', 'val': '空方契約金額'},
            '賣方交易': {'vol': '空方交易口數', 'val': '空方契約金額'},
            '賣出交易': {'vol': '空方交易口數', 'val': '空方契約金額'},
            '多空淨額交易': {'vol': '多空淨額交易口數', 'val': '多空淨額契約金額'},
            '買賣淨額交易': {'vol': '多空淨額交易口數', 'val': '多空淨額契約金額'},
            '交易淨額': {'vol': '多空淨額交易口數', 'val': '多空淨額契約金額'},
            '多方未平倉': {'vol': '多方未平倉口數', 'val': '多方未平倉契約金額'},
            '買方未平倉': {'vol': '多方未平倉口數', 'val': '多方未平倉契約金額'},
            '買進未平倉': {'vol': '多方未平倉口數', 'val': '多方未平倉契約金額'},
            '空方未平倉': {'vol': '空方未平倉口數', 'val': '空方未平倉契約金額'},
            '賣方未平倉': {'vol': '空方未平倉口數', 'val': '空方未平倉契約金額'},
            '賣出未平倉': {'vol': '空方未平倉口數', 'val': '空方未平倉契約金額'},
            '多空淨額未平倉': {'vol': '多空淨額未平倉口數', 'val': '多空淨額未平倉契約金額'},
            '買賣淨額未平倉': {'vol': '多空淨額未平倉口數', 'val': '多空淨額未平倉契約金額'},
            '未平倉淨額': {'vol': '多空淨額未平倉口數', 'val': '多空淨額未平倉契約金額'},
        }
        
        # 口數與金額的關鍵詞
        vol_keywords = ['口數', '數量', '口', '交易量', '成交量']
        val_keywords = ['金額', '價值', '契約金額', '契約價值']
        
        # 尋找表頭行
        header_row = None
        for row in table.find_all('tr'):
            cells = row.find_all(['th', 'td'])
            if not cells:
                continue
                
            row_text = ' '.join([cell.get_text(strip=True) for cell in cells])
            # 檢查是否包含表頭關鍵字
            if any(key in row_text for key in column_keys.keys()):
                header_row = row
                break
        
        if not header_row:
            return {}
            
        # 解析表頭
        result = {}
        header_cells = header_row.find_all(['th', 'td'])
        
        for i, cell in enumerate(header_cells):
            cell_text = cell.get_text(strip=True)
            
            # 檢查每個欄位標題
            for key, mapping in column_keys.items():
                if key in cell_text:
                    # 檢查是口數還是金額
                    if any(vk in cell_text for vk in vol_keywords):
                        result[mapping['vol']] = i
                    elif any(vk in cell_text for vk in val_keywords):
                        result[mapping['val']] = i
        
        return result


def parse_arguments():
    """解析命令行參數"""
    parser = argparse.ArgumentParser(description='台灣期貨交易所資料爬取工具')
    
    # 新增便捷的日期範圍參數
    parser.add_argument('--date-range', type=str, 
                        help='日期範圍: "today" 表示今天, "YYYY-MM-DD" 表示單日, "YYYY-MM-DD,YYYY-MM-DD" 表示範圍')
    
    # 原有的日期範圍參數（保持向後相容）
    parser.add_argument('--start_date', type=str, help='開始日期 (格式: YYYY/MM/DD)')
    parser.add_argument('--end_date', type=str, help='結束日期 (格式: YYYY/MM/DD)')
    parser.add_argument('--year', type=int, help='要爬取的年份，例如 2023')
    parser.add_argument('--month', type=int, help='要爬取的月份，例如 1-12')
    
    # 契約參數 - 支援逗號分隔的字串
    parser.add_argument('--contracts', type=str, default='ALL',
                        help='要爬取的契約代碼，例如 TX,TE,MTX，或使用 ALL 爬取所有契約')
    
    # 身份別參數
    parser.add_argument('--identities', type=str, nargs='+',
                        choices=IDENTITIES + ['ALL', 'NONE'], default=['ALL'],
                        help='要爬取的身份別，例如 自營商 投信 外資，或使用 ALL 爬取所有身份別，NONE 表示不爬取身份別資料')
    
    # 輸出參數
    parser.add_argument('--output_dir', type=str, default='output',
                        help='輸出目錄路徑')
    parser.add_argument('--filename', type=str, help='輸出檔名 (不含副檔名)')
    
    # 爬蟲配置參數
    parser.add_argument('--max_workers', type=int, default=10,
                        help='最大工作線程數')
    parser.add_argument('--delay', type=float, default=0.2,
                        help='請求間隔時間 (秒)')
    parser.add_argument('--max_retries', type=int, default=3,
                        help='最大重試次數')
    
    args = parser.parse_args()
    
    # 處理新的日期範圍參數
    if args.date_range:
        if args.date_range.lower() == 'today':
            # 今天
            today = datetime.datetime.now()
            start_date = today
            end_date = today
        elif ',' in args.date_range:
            # 日期範圍: YYYY-MM-DD,YYYY-MM-DD
            dates = args.date_range.split(',')
            start_date = datetime.datetime.strptime(dates[0].strip(), "%Y-%m-%d")
            end_date = datetime.datetime.strptime(dates[1].strip(), "%Y-%m-%d")
        else:
            # 單日: YYYY-MM-DD
            date = datetime.datetime.strptime(args.date_range, "%Y-%m-%d")
            start_date = date
            end_date = date
    elif args.year:
        # 如果指定了年份
        year = args.year
        if args.month:
            # 如果同時指定了月份，則爬取該月
            month = args.month
            start_date = datetime.datetime(year, month, 1)
            if month == 12:
                end_date = datetime.datetime(year + 1, 1, 1) - datetime.timedelta(days=1)
            else:
                end_date = datetime.datetime(year, month + 1, 1) - datetime.timedelta(days=1)
        else:
            # 只指定年份，爬取整年
            start_date = datetime.datetime(year, 1, 1)
            end_date = datetime.datetime(year, 12, 31)
    else:
        # 使用明確的開始和結束日期
        if not args.start_date:
            # 默認為當年初至今
            today = datetime.datetime.now()
            start_date = datetime.datetime(today.year, 1, 1)
            end_date = today
        else:
            # 解析用戶提供的日期
            start_date = datetime.datetime.strptime(args.start_date, "%Y/%m/%d")
            if args.end_date:
                end_date = datetime.datetime.strptime(args.end_date, "%Y/%m/%d")
            else:
                end_date = datetime.datetime.now()
    
    # 確保結束日期不超過今天
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if end_date > today:
        end_date = today
        logger.info(f"結束日期設定為今天: {today.strftime('%Y/%m/%d')}")
    
    args.start_date = start_date
    args.end_date = end_date
    
    # 處理契約邏輯 - 支援逗號分隔的字串
    if isinstance(args.contracts, str):
        if args.contracts.upper() == 'ALL':
            args.contracts = CONTRACTS
        else:
            args.contracts = [c.strip().upper() for c in args.contracts.split(',') if c.strip()]
    elif 'ALL' in args.contracts:
        args.contracts = CONTRACTS
    
    # 處理身份別邏輯
    if 'ALL' in args.identities:
        args.identities = IDENTITIES
    elif 'NONE' in args.identities:
        args.identities = None
    
    return args


def main():
    """主程序"""
    args = parse_arguments()
    
    # 顯示爬取配置
    logger.info("台灣期貨交易所資料爬取工具 (Python 版)")
    logger.info(f"爬取日期範圍: {args.start_date.strftime('%Y/%m/%d')} - {args.end_date.strftime('%Y/%m/%d')}")
    logger.info(f"契約: {', '.join(args.contracts)}")
    logger.info(f"身份別: {', '.join(args.identities) if args.identities else '不爬取身份別資料'}")
    
    # 初始化資料庫管理器
    if DB_AVAILABLE:
        db_manager = TaifexDatabaseManager()
        report_generator = DailyReportGenerator(db_manager)
        logger.info("資料庫系統已啟用")
    else:
        db_manager = None
        report_generator = None
    
    # 創建爬蟲實例
    crawler = TaifexCrawler(
        output_dir=args.output_dir,
        max_workers=args.max_workers,
        delay=args.delay,
        max_retries=args.max_retries
    )
    
    # 爬取資料
    df = crawler.crawl_date_range(
        args.start_date,
        args.end_date,
        contracts=args.contracts,
        identities=args.identities
    )
    
    # 保存資料
    if not df.empty:
        # 生成默認檔名
        if not args.filename:
            date_range = f"{args.start_date.strftime('%Y%m%d')}-{args.end_date.strftime('%Y%m%d')}"
            contracts_str = '_'.join(args.contracts)
            identity_str = '_'.join(args.identities) if args.identities else 'no_identity'
            args.filename = f"taifex_{date_range}_{contracts_str}_{identity_str}"
        
        # 1. 保存到傳統檔案格式
        csv_path, excel_path = crawler.save_data(df, args.filename)
        logger.info(f"已成功爬取 {len(df)} 筆資料")
        logger.info(f"CSV 檔案: {csv_path}")
        logger.info(f"Excel 檔案: {excel_path}")
        
        # 2. 保存到資料庫（如果可用）
        if db_manager and not df.empty:
            try:
                # 轉換資料格式以符合資料庫結構
                db_df = prepare_data_for_db(df)
                db_manager.insert_data(db_df)
                logger.info("資料已成功存入資料庫")
                
                # 生成30天日報（如果資料足夠）
                recent_data = db_manager.get_recent_data(30)
                if not recent_data.empty and len(recent_data) > 50:  # 確保有足夠資料
                    report = report_generator.generate_30day_report()
                    if report:
                        logger.info("30天分析報告已生成")
                
                # 匯出最新30天資料到固定檔案
                latest_30d_path = Path(args.output_dir) / "台期所最新30天資料.xlsx"
                db_manager.export_to_excel(latest_30d_path, days=30)
                logger.info(f"最新30天資料已匯出: {latest_30d_path}")
                
            except Exception as e:
                logger.error(f"資料庫操作失敗: {e}")
        
    else:
        logger.warning("未爬取到任何資料")
    
    logger.info("程式執行完成")


def prepare_data_for_db(df):
    """將爬蟲資料轉換為資料庫格式"""
    if df.empty:
        return pd.DataFrame()
    
    # 資料庫需要的欄位
    required_columns = [
        'date', 'contract_code', 'identity_type', 'position_type',
        'long_position', 'short_position', 'net_position'
    ]
    
    db_records = []
    
    for _, row in df.iterrows():
        base_record = {
            'date': row.get('交易日期', ''),
            'contract_code': row.get('契約', ''),
        }
        
        # 處理身份別資料
        if '身份別' in df.columns:
            base_record['identity_type'] = row.get('身份別', '')
        else:
            base_record['identity_type'] = '總計'
        
        # 處理多方部位
        if any(col for col in df.columns if '多方' in col and '口數' in col):
            long_cols = [col for col in df.columns if '多方' in col and '口數' in col]
            long_position = sum(row.get(col, 0) for col in long_cols)
            
            record_long = base_record.copy()
            record_long.update({
                'position_type': '多方',
                'long_position': long_position,
                'short_position': 0,
                'net_position': long_position
            })
            db_records.append(record_long)
        
        # 處理空方部位
        if any(col for col in df.columns if '空方' in col and '口數' in col):
            short_cols = [col for col in df.columns if '空方' in col and '口數' in col]
            short_position = sum(row.get(col, 0) for col in short_cols)
            
            record_short = base_record.copy()
            record_short.update({
                'position_type': '空方',
                'long_position': 0,
                'short_position': short_position,
                'net_position': -short_position
            })
            db_records.append(record_short)
        
        # 處理淨部位
        if any(col for col in df.columns if '淨部位' in col):
            net_cols = [col for col in df.columns if '淨部位' in col]
            net_position = sum(row.get(col, 0) for col in net_cols)
            
            record_net = base_record.copy()
            record_net.update({
                'position_type': '淨部位',
                'long_position': 0,
                'short_position': 0,
                'net_position': net_position
            })
            db_records.append(record_net)
    
    if not db_records:
        # 如果沒有識別到標準格式，創建基本記錄
        for _, row in df.iterrows():
            record = {
                'date': row.get('交易日期', ''),
                'contract_code': row.get('契約', ''),
                'identity_type': row.get('身份別', '總計'),
                'position_type': '未分類',
                'long_position': 0,
                'short_position': 0,
                'net_position': 0
            }
            db_records.append(record)
    
    return pd.DataFrame(db_records)


if __name__ == "__main__":
    main() 