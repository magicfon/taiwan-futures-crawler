@echo off
chcp 65001 > nul
echo 🚀 台期所交易量資料爬取 (下午2點資料)
echo ===========================================
echo 正在爬取今日交易量資料...
echo.

python taifex_crawler.py --date-range today --contracts TX,TE,MTX,ZMX,NQF --identities ALL --data_type TRADING --max_workers 5 --delay 1.0 --check_days 7

echo.
echo 爬取完成！資料已上傳到Google Sheets的「交易量資料」分頁
echo 💡 提示：請在下午3點半後執行 crawl_complete_data.bat 來爬取完整資料
pause 