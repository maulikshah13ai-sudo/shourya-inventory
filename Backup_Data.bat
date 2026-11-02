@echo off
title 1-Click Database Backup
color 0a

echo ========================================================
echo   BACKING UP BAGS & FOOTWEAR SHOP DATABASE
echo ========================================================
echo.

cd /d "%~dp0"
python backup_database.py

echo.
echo Backup completed. You can safely copy the backup file to a USB drive or Google Drive.
pause
