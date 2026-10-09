@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===============================================================
echo   Film Index - rebuild snapshot (Windows)
echo   (only needed for file mode; service mode updates automatically)
echo ===============================================================
echo.
python assets/worker.py --build
if errorlevel 1 py assets/worker.py --build
echo.
pause
