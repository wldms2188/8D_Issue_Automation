@echo off
cd /d "%~dp0"
py app\main_v29.py
if errorlevel 1 python app\main_v29.py
pause
