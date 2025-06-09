@echo off
chcp 65001 > nul
echo 📈 台期所完整資料爬取 (下午3點半資料)
echo ===========================================
echo 正在爬取今日完整資料 (包含交易量和未平倉)...
echo.

python taifex_crawler.py --date-range today --contracts TX,TE,MTX --identities ALL --data_type COMPLETE --max_workers 5 --delay 1.0

echo.
echo 爬取完成！完整資料已上傳到Google Sheets的「完整資料」分頁
echo 💡 提示：完整資料包含交易量和未平倉部位資訊
pause 