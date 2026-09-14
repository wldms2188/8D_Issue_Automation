@echo off
cd /d "%~dp0"
echo ========================================
echo 8D Issue Automation v3.0.4
echo ========================================
echo.
py app\main_v304.py
if errorlevel 1 (
    echo.
    echo Python launcher 'py' failed. Trying 'python'...
    python app\main_v304.py
)
echo.
echo ========================================
echo Program finished.
echo ========================================
pause
