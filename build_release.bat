@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo =====================================================
echo   8D Issue Automation - Stable Installer Build
echo   Application baseline: 4ea1958
echo =====================================================

py -3 -m pip install --upgrade pip
if errorlevel 1 goto :fail
py -3 -m pip install -r requirements.txt
if errorlevel 1 goto :fail

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if not exist installer\output mkdir installer\output

echo.
echo [1/2] Building current enterprise EXE...
py -3 -m PyInstaller --noconfirm --clean 8D_Issue_Automation.spec
if errorlevel 1 goto :fail

if not exist "dist\8D_Issue_Automation.exe" (
  echo EXE was not created.
  goto :fail
)

echo.
echo [2/2] Building Inno Setup installer...

set "ISCC="
where iscc >nul 2>nul
if not errorlevel 1 set "ISCC=iscc"

if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not defined ISCC (
  echo.
  echo Inno Setup 6 was not found.
  echo Install Inno Setup 6, then run this file again.
  echo The EXE itself is already available at:
  echo   dist\8D_Issue_Automation.exe
  goto :done
)

"%ISCC%" "installer\8D_Issue_Automation.iss"
if errorlevel 1 goto :fail

if not exist "installer\output\8D_Issue_Automation_Stable_4ea1958_Setup.exe" (
  echo Installer output was not created.
  goto :fail
)

echo.
echo Installer complete:
echo   installer\output\8D_Issue_Automation_Stable_4ea1958_Setup.exe
goto :done

:fail
echo.
echo BUILD FAILED. Check the messages above.
pause
exit /b 1

:done
echo.
echo Release build finished.
pause
