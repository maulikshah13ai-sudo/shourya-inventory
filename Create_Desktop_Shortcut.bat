@echo off
setlocal
echo Creating Desktop Shortcut for Bags & Shoes Inventory System...

set "TARGET_DIR=%~dp0"
set "BAT_FILE=%TARGET_DIR%Start_Shop_App.bat"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\Shop Inventory System.lnk"

powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath='%BAT_FILE%'; $s.WorkingDirectory='%TARGET_DIR%'; $s.Description='Launch Bags & Shoes Inventory App'; $s.Save()"

if exist "%SHORTCUT_PATH%" (
    echo.
    echo [SUCCESS] Shortcut created on your Desktop: "Shop Inventory System"
    echo You can now double-click it directly from your desktop anytime!
) else (
    echo Could not create desktop shortcut automatically. You can double click Start_Shop_App.bat directly in this folder.
)
echo.
pause
