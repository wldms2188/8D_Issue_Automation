@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [8D Issue Automation v3.2.0 RECOVERY STEP9]
py -3 app\main_recovery_step9.py
if errorlevel 1 py app\main_recovery_step9.py
if errorlevel 1 pause
