@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    start "" http://127.0.0.1:8766/
    py dashboard_server.py
    goto :eof
)
where python >nul 2>nul
if %errorlevel%==0 (
    start "" http://127.0.0.1:8766/
    python dashboard_server.py
    goto :eof
)
echo Python is required to run the dashboard.
echo Install Python and then run this file again.
pause
