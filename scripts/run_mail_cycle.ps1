$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = "python"
$syncScript = Join-Path $repoRoot "scripts\sync_outlook_mail.py"
$logDir = Join-Path $repoRoot "logs"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $logDir "mail_cycle_$timestamp.log"

New-Item -ItemType Directory -Force -Path $logDir | Out-Null

& $python $syncScript --process-after-sync 2>&1 | Tee-Object -FilePath $logPath
if ($LASTEXITCODE -ne 0) {
    throw "Mail cycle failed. See $logPath"
}

Write-Host "Mail cycle completed successfully. Log: $logPath"
