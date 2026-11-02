@echo off
:: Batch script to allow inbound access on port 5000 for shop phones/tablets
echo Requesting administrator privileges to allow local Wi-Fi connections...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Adding Windows Firewall rule for Port 5000...
    netsh advfirewall firewall add rule name="Bags Shoes Shop App (Port 5000)" dir=in action=allow protocol=TCP localport=5000
    echo.
    echo [SUCCESS] Windows Firewall is now open for port 5000!
    echo Phones, tablets, and laptops on your shop Wi-Fi can now connect smoothly.
) else (
    echo.
    echo [NOTE] Please right-click this file and choose "Run as administrator" to apply the firewall rule.
)
echo.
pause
