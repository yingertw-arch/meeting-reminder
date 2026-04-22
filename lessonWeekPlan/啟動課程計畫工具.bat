@echo off
chcp 65001 >nul
title 課程計畫生成工具
echo ================================================
echo   海山國小 課程計畫生成工具
echo ================================================
echo.
echo 正在啟動伺服器，請稍候...
echo.
echo 啟動後請開啟瀏覽器前往：http://localhost:5001
echo 關閉此視窗即可停止服務。
echo.
cd /d "%~dp0"
python server.py
pause
