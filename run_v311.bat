@echo off
cd /d "%~dp0"
py app\main_v311.py
if errorlevel 1 pause
