@echo off
cd /d "%~dp0"
py app\run_v27.py
if errorlevel 1 python app\run_v27.py
pause
