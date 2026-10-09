@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================================
echo   Film Index - local server (Windows)
echo ===============================================================
echo.
python assets/worker.py --serve
if errorlevel 1 py assets/worker.py --serve
pause
