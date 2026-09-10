@echo off
cd /d "%~dp0"
py app\main.py
if errorlevel 1 python app\main.py
pause
