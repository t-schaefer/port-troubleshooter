<#
.SYNOPSIS
  Removes the Port Troubleshooter Scheduled Task and its firewall rule(s).
  Run from an elevated PowerShell:  powershell -ExecutionPolicy Bypass -File .\uninstall-task-windows.ps1
  data\, config.json, switches.json and cv_token.txt stay where they are.
#>
$TaskName = "PortTroubleshooterService"
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Task '$TaskName' removed."
}
Get-NetFirewallRule -DisplayName "Port Troubleshooter (TCP *)" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
Write-Host "Done."
