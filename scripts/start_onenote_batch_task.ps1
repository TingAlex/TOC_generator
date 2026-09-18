param(
    [Parameter(Mandatory = $true)]
    [string]$Plan,

    [string]$TaskName = "TOCGenerator-OneNoteBatch",

    [ValidateRange(1, 20)]
    [int]$MaxAttempts = 3,

    [ValidateRange(0, 3600)]
    [int]$RetryDelay = 30
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    throw "Project Python not found: $python. Run uv sync first."
}

$planPath = (Resolve-Path -LiteralPath $Plan).Path
# Quote only the resolved path. All other arguments are generated here.
$quotedPlan = [char]34 + $planPath + [char]34
$arguments = "-m tocgen.cli.onenote_batch --plan $quotedPlan --write --max-attempts $MaxAttempts --retry-delay $RetryDelay"

$action = New-ScheduledTaskAction `
    -Execute $python `
    -Argument $arguments `
    -WorkingDirectory $repoRoot
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(10)
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 12) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "TOC_generator native OneNote page copy with automatic restart." `
    -Force | Out-Null
Start-ScheduledTask -TaskName $TaskName

Write-Output "Started Windows scheduled task: $TaskName"
Write-Output "Plan: $planPath"
Write-Output "Status: Get-ScheduledTask -TaskName '$TaskName' | Get-ScheduledTaskInfo"
