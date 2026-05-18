@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: ── Прокси из config.env ──────────────────────────────────────────────────────
set PIP_PROXY=
if exist "%~dp0config.env" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%~dp0config.env") do (
        if /i "%%A"=="PIP_PROXY" set PIP_PROXY=%%B
    )
)

if not "%PIP_PROXY%"=="" (
    set PIP_PROXY_FLAG=--proxy "%PIP_PROXY%"
) else (
    set PIP_PROXY_FLAG=
)

:: ── Зависимости ──────────────────────────────────────────────────────────────
echo.
echo [1/2] Проверка зависимостей...
pip install -r requirements.txt --quiet %PIP_PROXY_FLAG% 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ОШИБКА при установке пакетов.
    echo  Проверьте значение PIP_PROXY в config.env.
    pause
    exit /b 1
)

:: ── Сервер ───────────────────────────────────────────────────────────────────
echo [2/2] Запуск Flask-сервера...
start "TimeSheet Auto — server" cmd /k python app.py

set /a attempts=0
:check_server
set /a attempts+=1
if %attempts% gtr 30 (
    echo.
    echo  Сервер не запустился за 30 секунд.
    echo  Смотрите окно "TimeSheet Auto — server" для деталей ошибки.
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
