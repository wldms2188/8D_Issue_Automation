@echo off
cd /d "%~dp0"
py app\main_v295.py
if errorlevel 1 python app\main_v295.py
pause
