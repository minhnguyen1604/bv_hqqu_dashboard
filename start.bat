@echo off
title Khoi dong Baoviet HQQU Dashboard
echo ===================================================
echo     KHOI DONG BAOVIET HQQU DASHBOARD (PORT 8000)
echo ===================================================
echo.
echo Dang khoi dong may chu local...
start /b python -m http.server 8000
timeout /t 2 >nul
echo.
echo Dang tu dong mo trinh duyet...
start http://localhost:8000/index.html
echo.
echo May chu dang chay tai: http://localhost:8000/index.html
echo De tat may chu, hay dong cua so nay.
echo.
pause
