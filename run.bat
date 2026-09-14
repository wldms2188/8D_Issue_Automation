@echo off
cd /d "%~dp0"
py app\main_v300.py
if errorlevel 1 python app\main_v300.py
pause
