@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] Installing / checking dependencies...
pip install -r requirements.txt --quiet --break-system-packages 2>nul ^
  || pip install -r requirements.txt --quiet

echo [2/2] Starting Flask server...
:: cmd /k keeps the window open so errors are visible
start "TimeSheet Auto — server" cmd /k python app.py

:: Wait until port 5000 is actually listening (up to 30 s)
set /a attempts=0
:check_server
set /a attempts+=1
if %attempts% gtr 30 (
    echo.
    echo ERROR: server did not start in 30 seconds.
    echo Check the "TimeSheet Auto" window for the Python error.
    pause
    exit /b 1
)
powershell -noprofile -command ^
  "try{(New-Object Net.Sockets.TcpClient).Connect('127.0.0.1',5000);exit 0}catch{exit 1}" 2>nul
if %errorlevel% neq 0 (
    timeout /t 1 /nobreak >nul
    goto check_server
)

start http://localhost:5000
