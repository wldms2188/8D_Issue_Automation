@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [8D Issue Automation 1.0 - Enterprise English Direct Analysis]
echo Starting enterprise UI...

py -3 app\main_enterprise_final.py
if errorlevel 1 py app\main_enterprise_final.py
if errorlevel 1 (
  echo.
  echo Program ended with an error.
  pause
)
