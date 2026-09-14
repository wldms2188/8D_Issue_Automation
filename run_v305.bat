@echo off
cd /d "%~dp0"
echo ========================================
echo 8D Issue Automation v3.0.5
echo ========================================
echo.
py app\main_v305.py
if errorlevel 1 (
    echo.
    echo Python launcher 'py' failed. Trying 'python'...
    python app\main_v305.py
)
echo.
echo ========================================
echo Program finished.
echo ========================================
pause
