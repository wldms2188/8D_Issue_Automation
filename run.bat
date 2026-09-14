@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [8D Issue Automation STEP14 FIX4]
py -3 app\main_recovery_step14_fix2.py
if errorlevel 1 py app\main_recovery_step14_fix2.py
if errorlevel 1 pause
