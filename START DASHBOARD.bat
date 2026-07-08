@echo off
title Modicare GSTR-1 Dashboard
echo.
echo  ==========================================
echo   Modicare GSTR-1 Dashboard - Asc Global Ai
echo  ==========================================
echo.
echo  Starting dashboard, please wait...
echo  Browser will open automatically.
echo.
cd /d "%~dp0"
C:\Users\Naresh\AppData\Local\Programs\Python\Python312\python.exe -m streamlit run gstr1_dashboard.py --server.port 8502 --server.headless false
pause
