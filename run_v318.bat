@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [8D Issue Automation v3.1.8]
py -3 app\main_v318.py
if errorlevel 1 (
  echo.
  echo Program ended with an error.
  pause
)
