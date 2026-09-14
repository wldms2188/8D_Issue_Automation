@echo off
cd /d "%~dp0"
echo ========================================
echo 8D Issue Automation v3.0.9 DEBUG
 echo ========================================
echo.
if exist error_v309.log del /q error_v309.log
py app\main_v309_debug.py
if errorlevel 1 (
    echo.
    echo Python launcher failed. Trying python...
    python app\main_v309_debug.py
)
echo.
echo ========================================
echo Diagnostic log: error_v309.log
 echo ========================================
if exist error_v309.log (
    echo.
    type error_v309.log
) else (
    echo No exception log was created.
)
echo.
pause
