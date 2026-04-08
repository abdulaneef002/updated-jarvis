$ErrorActionPreference = "Stop"

$taskName = "Project_JARVIS_AutoStart"
$projectRoot = $PSScriptRoot
$scriptPath = Join-Path $projectRoot "start_jarvis_background.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Error "Startup launcher not found at $scriptPath"
}

$powerShellExe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
$currentUser = "$env:USERDOMAIN\$env:USERNAME"

$action = New-ScheduledTaskAction -Execute $powerShellExe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Starts Project JARVIS automatically at user logon." -Force | Out-Null
Write-Output "Installed scheduled task: $taskName"
Write-Output "JARVIS will auto-start at login for user $currentUser"
