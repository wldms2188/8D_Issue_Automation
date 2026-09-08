@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
py app\main.py
if errorlevel 1 python app\main.py
pause
