@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [8D Issue Automation 1.0 FINAL]
py -3 app\main_final.py
if errorlevel 1 py app\main_final.py
if errorlevel 1 pause
