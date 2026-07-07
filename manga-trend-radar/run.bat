@echo off
REM manga-trend-radar 起動バッチ（タスクスケジューラから呼ぶ）
chcp 65001 > nul
cd /d "%~dp0"
python main.py
