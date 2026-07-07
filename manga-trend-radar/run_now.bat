@echo off
REM 手動実行用（ダブルクリックでOK）：曜日に関係なく収集→選抜→Google Chat送信
chcp 65001 > nul
cd /d "%~dp0"
python main.py --force
echo.
pause
