@echo off
rem OpsWatch daily health check (10:00)
cd /d C:\Users\ssasa\tools\ops-watch
python ops_watch.py
exit /b %ERRORLEVEL%
