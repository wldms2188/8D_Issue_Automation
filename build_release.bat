@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo   8D Issue Automation 1.0 - Enterprise Release
 echo ================================================

py -3 -m pip install --upgrade pip
if errorlevel 1 goto :fail
py -3 -m pip install pyinstaller python-pptx openpyxl pillow tkinterdnd2
if errorlevel 1 goto :fail

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

py -3 -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "8D_Issue_Automation" ^
  --paths app ^
  --collect-all pptx ^
  --collect-all openpyxl ^
  --collect-all tkinterdnd2 ^
  --hidden-import PIL ^
  app\main_enterprise.py
if errorlevel 1 goto :fail

echo.
echo EXE build complete: dist\8D_Issue_Automation.exe

where iscc >nul 2>nul
if errorlevel 1 (
  echo.
  echo Inno Setup is not installed or ISCC is not in PATH.
  echo EXE is ready. To create Setup.exe, install Inno Setup 6 and run:
  echo   iscc installer\8D_Issue_Automation.iss
  goto :done
)

if not exist installer\output mkdir installer\output
iscc installer\8D_Issue_Automation.iss
if errorlevel 1 goto :fail

echo.
echo Installer complete: installer\output\8D_Issue_Automation_Setup.exe
goto :done

:fail
echo.
echo BUILD FAILED. See the error above.
pause
exit /b 1

:done
echo.
echo Release build finished.
pause
