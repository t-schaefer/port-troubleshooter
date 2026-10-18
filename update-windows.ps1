<#
.SYNOPSIS
  Updates the Port Troubleshooter and restarts its task (PortTroubleshooterService):
  stops the task (and a leftover python running this server.py), updates the files -
  `git pull --ff-only` in a git clone, otherwise GitHub's main.zip copied over (nothing
  deleted) - and starts the task again. Usually started by double-clicking
  Update-PortTroubleshooter.bat (self-elevates).
  data\, config.json, switches.json, future.json, cv_token.txt and service.log are not in
  git / the zip and are never touched.
#>
$ErrorActionPreference = "Stop"
$TaskName = "PortTroubleshooterService"
$RepoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ServerPy = Join-Path $RepoDir "server.py"
$ZipUrl = "https://github.com/t-schaefer/port-troubleshooter/archive/refs/heads/main.zip"
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
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($git -and (Test-Path (Join-Path $RepoDir ".git"))) {
        Write-Host "git pull --ff-only ..."
        # safe.directory: also works when the clone belongs to another user (e.g. run as SYSTEM from the landing page)
        git -c "safe.directory=$($RepoDir -replace '\\','/')" pull --ff-only
        if ($LASTEXITCODE -ne 0) { throw "git pull failed (local changes or diverged branch?)" }
    } else {
        Write-Host "No git clone here - downloading $ZipUrl ..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $tmp = Join-Path $env:TEMP ("porttrouble-update-" + [guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $tmp | Out-Null
        $zip = Join-Path $tmp "main.zip"
        Invoke-WebRequest -Uri $ZipUrl -OutFile $zip -UseBasicParsing
        Expand-Archive -Path $zip -DestinationPath $tmp -Force
        $src = Get-ChildItem -Path $tmp -Directory | Select-Object -First 1
        robocopy $src.FullName $RepoDir /E /NFL /NDL /NJH /NJS /NP | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "copying the downloaded files failed (robocopy $LASTEXITCODE)" }
        Remove-Item -Recurse -Force $tmp
    }
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
