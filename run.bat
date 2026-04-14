@echo off
cd /d "%~dp0"
start "TimeSheet Auto" python app.py
timeout /t 2 /nobreak >nul
start http://localhost:5000
