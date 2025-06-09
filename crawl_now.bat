@echo off
chcp 65001 > nul
echo ⚡ 台期所智能爬取 (自動判斷資料類型)
echo ===========================================
echo 系統將根據當前時間自動判斷要爬取的資料類型：
echo   - 14:00-15:00: 爬取交易量資料
echo   - 其他時間: 爬取完整資料
echo.

python daily_crawler_schedule.py --mode now

pause 