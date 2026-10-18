@echo off
rem Runs the Port Troubleshooter in this window (no service) - for a quick test.
rem     Start-PortTroubleshooter.bat        -> http://localhost:8098/
setlocal
set "HERE=%~dp0"
set "PORT=%~1"
if "%PORT%"=="" set "PORT=8098"
cd /d "%HERE%"
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python was not found on PATH - install Python 3.8+ and tick "Add python.exe to PATH".
    pause
    exit /b 1
)
echo Port Troubleshooter - http://localhost:%PORT%/   Ctrl+C stops it.
python server.py --port %PORT%
echo.
pause
