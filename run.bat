@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [8D Issue Automation 1.0 - Enterprise English Direct Analysis]
echo Checking drag and drop support...

py -3 -c "import tkinterdnd2" >nul 2>&1
if errorlevel 1 (
  echo Installing tkinterdnd2 for Drag ^& Drop...
  py -3 -m pip install tkinterdnd2
)

echo Starting enterprise UI...
py -3 app\main_enterprise_final.py
if errorlevel 1 py app\main_enterprise_final.py
if errorlevel 1 (
  echo.
  echo Program ended with an error.
  pause
)
