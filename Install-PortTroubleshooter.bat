@echo off
rem One double-click install of the Port Troubleshooter as a Windows service (Scheduled Task
rem as SYSTEM, port 8092). Other port: Install-PortTroubleshooter.bat 8110
rem Self-elevates via UAC, pauses at the end. Updates: Update-PortTroubleshooter.bat
setlocal
set "HERE=%~dp0"
set "PORT=%~1"
if "%PORT%"=="" set "PORT=8092"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator rights...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%PORT%' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%install-task-windows.ps1" -Port %PORT%
echo.
pause
