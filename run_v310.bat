@echo off
cd /d "%~dp0"
py app\main_v310.py
if errorlevel 1 pause
