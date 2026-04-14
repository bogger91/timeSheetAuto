@echo off
cd /d "%~dp0"
echo Запуск сервера...
start "TimeSheet Auto" python app.py

:: Ждём, пока Flask реально откроет порт 5000 (до 30 секунд)
set /a attempts=0
:check_server
set /a attempts+=1
if %attempts% gtr 30 (
    echo.
    echo Сервер не запустился за 30 секунд.
    echo Проверьте окно "TimeSheet Auto" — там должна быть ошибка.
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
