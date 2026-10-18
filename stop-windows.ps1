<#
.SYNOPSIS
  Stops the Port Troubleshooter service: the Scheduled Task "PortTroubleshooterService" (waits until it
  is stopped) and any python still running server.py.
  It starts again at the next boot, with Update-PortTroubleshooter.bat,
  or by hand:  schtasks /run /tn PortTroubleshooterService
  Usually started by double-clicking Stop-PortTroubleshooter.bat (self-elevates).
#>
$ErrorActionPreference = "Stop"
$TaskName = "PortTroubleshooterService"
$ServerPy = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "server.py"

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Stopping task '$TaskName' ..."
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt 30; $i++) {
        if ((Get-ScheduledTask -TaskName $TaskName).State -ne "Running") { break }
        Start-Sleep -Seconds 1
    }
    Write-Host "Task state: $((Get-ScheduledTask -TaskName $TaskName).State)"
} else {
    Write-Host "Task '$TaskName' is not installed." -ForegroundColor Yellow
}
Get-CimInstance Win32_Process -Filter "Name like 'python%'" |
    Where-Object { $_.CommandLine -and $_.CommandLine -like "*$ServerPy*" } |
    ForEach-Object {
        Write-Host "Stopping python (PID $($_.ProcessId)) running $ServerPy"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
Write-Host "Stopped. Start again: schtasks /run /tn $TaskName  (or Update-PortTroubleshooter.bat)" -ForegroundColor Green
