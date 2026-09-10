@echo off
cd /d "%~dp0"
py app\run_v25.py
if errorlevel 1 python app\run_v25.py
pause
