@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
python lead_finder.py %*
echo.
echo Готово. Excel лежить у папці output
pause
