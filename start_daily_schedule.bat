@echo off
chcp 65001 > nul
echo 🤖 台期所每日自動爬取排程器
echo ===========================================
echo 排程時間：
echo   - 週一到週五 14:00: 爬取交易量資料
echo   - 週一到週五 15:30: 爬取完整資料
echo.
echo 正在啟動排程器...
echo 按 Ctrl+C 可停止排程器
echo.

python daily_crawler_schedule.py --mode schedule

pause 