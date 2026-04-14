@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: ── Прокси ──────────────────────────────────────────────────────────────────
:: Сохраняем прокси в proxy.cfg рядом со скриптом, чтобы не вводить каждый раз.
set PROXY_CFG=%~dp0proxy.cfg
set PIP_PROXY=

if exist "%PROXY_CFG%" (
    set /p PIP_PROXY=<"%PROXY_CFG%"
)

if "%PIP_PROXY%"=="" (
    echo.
    echo  Прокси для установки пакетов не задан.
    echo  Введите адрес прокси и нажмите Enter.
    echo  Формат:  http://proxy.company.ru:3128
    echo  Или оставьте пустым, если прокси не нужен.
    echo.
    set /p PIP_PROXY=  Прокси:
    echo.

    if not "%PIP_PROXY%"=="" (
        echo %PIP_PROXY%>"%PROXY_CFG%"
        echo  Прокси сохранён в proxy.cfg
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
    echo  Если прокси неверный, удалите файл proxy.cfg и запустите снова.
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
