@echo off
cd /d "%~dp0"
py app\run_v26.py
if errorlevel 1 python app\run_v26.py
pause
