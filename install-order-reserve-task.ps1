$ErrorActionPreference = "Stop"
$taskName = "Tharaldsen Dashboard - Ordrereserve"
$scriptPath = Join-Path $PSScriptRoot "update-order-reserve.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument ('-NoProfile -ExecutionPolicy Bypass -File "' + $scriptPath + '"')
$trigger = New-ScheduledTaskTrigger -Daily -At 07:00
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

try {
	Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description "Oppdaterer dashboardets ordrereserve fra Excel K7 hver dag kl. 07:00." -Force | Out-Null
	Write-Output "Oppgaven er registrert: $taskName"
} catch {
	Write-Error "Kunne ikke registrere oppgaven. Kjor dette scriptet som administrator."
	exit 1
}
