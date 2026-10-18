@echo off
rem One double-click update of the Port Troubleshooter service. Self-elevates via UAC, pauses at the end.
setlocal
set "HERE=%~dp0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator rights...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%update-windows.ps1"
echo.
pause
