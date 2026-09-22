@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo   8D Issue Automation - EXE Build
echo   Stable baseline: 4ea1958
echo ================================================

py -3 -m pip install --upgrade pip
if errorlevel 1 goto :fail
py -3 -m pip install -r requirements.txt
if errorlevel 1 goto :fail

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

py -3 -m PyInstaller --noconfirm --clean 8D_Issue_Automation.spec
if errorlevel 1 goto :fail

if not exist "dist\8D_Issue_Automation.exe" goto :fail

echo.
echo EXE build complete:
echo   dist\8D_Issue_Automation.exe
goto :done

:fail
echo.
echo EXE BUILD FAILED. Check the messages above.
pause
exit /b 1

:done
echo.
pause
