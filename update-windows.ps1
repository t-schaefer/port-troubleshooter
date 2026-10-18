<#
.SYNOPSIS
  Updates the Port Troubleshooter and restarts its task (PortTroubleshooterService):
  stops the task (and a leftover python running this server.py), `git pull --ff-only`,
  starts the task again. Usually started by double-clicking Update-PortTroubleshooter.bat.
  data\, config.json, switches.json, cv_token.txt and service.log are not in git and are never touched.
#>
$ErrorActionPreference = "Stop"
$TaskName = "PortTroubleshooterService"
$RepoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ServerPy = Join-Path $RepoDir "server.py"
Set-Location $RepoDir
Write-Host "Folder: $RepoDir"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Write-Host "Stopping task '$TaskName' ..."
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt 30; $i++) {
        if ((Get-ScheduledTask -TaskName $TaskName).State -ne "Running") { break }
        Start-Sleep -Seconds 1
    }
} else {
    Write-Host "Task '$TaskName' is not installed - only updating the files (Install-PortTroubleshooter.bat installs it)." -ForegroundColor Yellow
}
Get-CimInstance Win32_Process -Filter "Name like 'python%'" |
    Where-Object { $_.CommandLine -and $_.CommandLine -like "*$ServerPy*" } |
    ForEach-Object { Write-Host "Stopping leftover python (PID $($_.ProcessId))"; Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

$updated = $false
try {
    if (-not (Test-Path (Join-Path $RepoDir ".git"))) { throw "this folder is not a git clone" }
    Write-Host "git pull --ff-only ..."
    git -c "safe.directory=$($RepoDir -replace '\\','/')" pull --ff-only
    if ($LASTEXITCODE -ne 0) { throw "git pull failed (local changes or diverged branch?)" }
    $updated = $true
    Write-Host "Files updated." -ForegroundColor Green
} catch {
    Write-Host "UPDATE FAILED: $_" -ForegroundColor Red
    Write-Host "The task is started again with the files as they were." -ForegroundColor Yellow
}

if ($task) {
    Start-ScheduledTask -TaskName $TaskName
    Start-Sleep -Seconds 3
    Write-Host "Task state: $((Get-ScheduledTask -TaskName $TaskName).State)"
}
if (-not $updated) { exit 1 }
Write-Host "Done."
