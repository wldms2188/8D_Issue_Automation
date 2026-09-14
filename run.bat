@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [8D Issue Automation - latest launcher]
echo Checking latest v3.2.5 branch...

git fetch origin test/v308-synthetic
if errorlevel 1 goto :giterror

git switch test/v308-synthetic
if errorlevel 1 goto :giterror

git pull
if errorlevel 1 goto :giterror

echo [8D Issue Automation v3.2.5]
py -3 app\main_v325.py
if errorlevel 1 py app\main_v325.py
if errorlevel 1 goto :runerror
goto :eof

:giterror
echo.
echo Git branch update failed. Please check local changes or network connection.
pause
goto :eof

:runerror
echo.
echo Program ended with an error.
pause
