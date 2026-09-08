@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean --onefile --windowed --name 8D_이슈_자동화 app\main.py
pause
