@echo off
title Bags & Shoes Inventory System (Multi-User Production Server)
color 0b

echo ====================================================================
echo   APEX BAGS & FOOTWEAR WHOLESALE + RETAIL INVENTORY SYSTEM
echo ====================================================================
echo.
echo   Starting Production Multi-User Server (Waitress)...
echo   Please wait a moment while the server launches.
echo.

cd /d "%~dp0"

REM Launch default browser after 2 seconds
start "" timeout /t 2 /nobreak >nul & start http://localhost:5000

REM Run production server
python run_production.py

pause
