param(
    [string]$SyncArgs = "--process-after-sync"
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$repoRoot = Split-Path -Parent $PSScriptRoot
$powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$python = "python"
$syncScript = Join-Path $repoRoot "scripts\sync_outlook_mail.py"
$dataDir = Join-Path $repoRoot "data"
$logDir = Join-Path $repoRoot "logs"
$screenshotsRoot = Join-Path $logDir "screenshots"
$timingPath = Join-Path $dataDir "mail_cycle_timings.json"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $logDir "mail_cycle_$timestamp.log"
$screenshotDir = Join-Path $screenshotsRoot $timestamp
$screenshotIntervalSeconds = 20
$stdoutPath = Join-Path $logDir "mail_cycle_$timestamp.stdout.log"
$stderrPath = Join-Path $logDir "mail_cycle_$timestamp.stderr.log"
$exitCodePath = Join-Path $logDir "mail_cycle_$timestamp.exitcode.txt"
$removedScreenshotDirs = @()

New-Item -ItemType Directory -Force -Path $logDir | Out-Null
New-Item -ItemType Directory -Force -Path $screenshotsRoot | Out-Null
New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
Get-ChildItem -Path $screenshotsRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $directory = $_
    $folderDate = $null
    if ($directory.Name -match '^(?<date>\d{8})_\d{6}$') {
        $folderDate = [datetime]::ParseExact($Matches['date'], 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    else {
        $folderDate = $directory.LastWriteTime.Date
    }

    if ($folderDate.Date -lt (Get-Date).Date) {
        Remove-Item -LiteralPath $directory.FullName -Recurse -Force
        $removedScreenshotDirs += $directory.FullName
    }
}
New-Item -ItemType Directory -Force -Path $screenshotDir | Out-Null
New-Item -ItemType File -Force -Path $stdoutPath | Out-Null
New-Item -ItemType File -Force -Path $stderrPath | Out-Null
if (Test-Path $exitCodePath) {
    Remove-Item -LiteralPath $exitCodePath -Force
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Crimson Automation - Mail Cycle Monitor"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1040, 1060)
$form.MinimumSize = New-Object System.Drawing.Size(980, 1020)
$form.BackColor = [System.Drawing.Color]::FromArgb(18, 3, 5)
$form.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Mail Cycle Monitor"
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 20, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(24, 18)
$title.Size = New-Object System.Drawing.Size(500, 42)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Outlook desktop -> raw/mail -> catch-up unread -> SQLite -> attachments -> classify -> decision queue"
$subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)
$subtitle.Location = New-Object System.Drawing.Point(26, 58)
$subtitle.Size = New-Object System.Drawing.Size(880, 24)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: starting"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$statusLabel.Location = New-Object System.Drawing.Point(26, 102)
$statusLabel.Size = New-Object System.Drawing.Size(360, 24)

$purposeLabel = New-Object System.Windows.Forms.Label
$purposeLabel.Text = "Purpose: sync mail and refresh the Needs Decision queue"
$purposeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$purposeLabel.Location = New-Object System.Drawing.Point(26, 130)
$purposeLabel.Size = New-Object System.Drawing.Size(560, 24)

$elapsedLabel = New-Object System.Windows.Forms.Label
$elapsedLabel.Text = "Elapsed: 00:00:00"
$elapsedLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$elapsedLabel.Location = New-Object System.Drawing.Point(26, 158)
$elapsedLabel.Size = New-Object System.Drawing.Size(220, 24)

$currentStageLabel = New-Object System.Windows.Forms.Label
$currentStageLabel.Text = "Current stage: waiting to start"
$currentStageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$currentStageLabel.Location = New-Object System.Drawing.Point(26, 186)
$currentStageLabel.Size = New-Object System.Drawing.Size(540, 24)

$stageActivityLabel = New-Object System.Windows.Forms.Label
$stageActivityLabel.Text = "Stage activity: waiting for the first stage"
$stageActivityLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$stageActivityLabel.Location = New-Object System.Drawing.Point(26, 214)
$stageActivityLabel.Size = New-Object System.Drawing.Size(620, 24)
$stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)

$childStateLabel = New-Object System.Windows.Forms.Label
$childStateLabel.Text = "Child process: not started"
$childStateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$childStateLabel.Location = New-Object System.Drawing.Point(620, 102)
$childStateLabel.Size = New-Object System.Drawing.Size(320, 24)
$childStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)

$pipelineStateLabel = New-Object System.Windows.Forms.Label
$pipelineStateLabel.Text = "Pipeline state: STOPPED"
$pipelineStateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$pipelineStateLabel.Location = New-Object System.Drawing.Point(620, 130)
$pipelineStateLabel.Size = New-Object System.Drawing.Size(340, 24)
$pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)

$activityStateLabel = New-Object System.Windows.Forms.Label
$activityStateLabel.Text = "Activity: waiting"
$activityStateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$activityStateLabel.Location = New-Object System.Drawing.Point(620, 158)
$activityStateLabel.Size = New-Object System.Drawing.Size(340, 24)
$activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)

$lastEventLabel = New-Object System.Windows.Forms.Label
$lastEventLabel.Text = "Last event: none yet"
$lastEventLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lastEventLabel.Location = New-Object System.Drawing.Point(620, 186)
$lastEventLabel.Size = New-Object System.Drawing.Size(360, 24)
$lastEventLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)

$currentStageEtaLabel = New-Object System.Windows.Forms.Label
$currentStageEtaLabel.Text = "Current stage ETA: n/a"
$currentStageEtaLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$currentStageEtaLabel.Location = New-Object System.Drawing.Point(620, 214)
$currentStageEtaLabel.Size = New-Object System.Drawing.Size(340, 24)
$currentStageEtaLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)

$nextStageEtaLabel = New-Object System.Windows.Forms.Label
$nextStageEtaLabel.Text = "Next stage ETA: n/a"
$nextStageEtaLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$nextStageEtaLabel.Location = New-Object System.Drawing.Point(620, 242)
$nextStageEtaLabel.Size = New-Object System.Drawing.Size(340, 24)
$nextStageEtaLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)

$remainingEtaLabel = New-Object System.Windows.Forms.Label
$remainingEtaLabel.Text = "Remaining ETA: n/a"
$remainingEtaLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$remainingEtaLabel.Location = New-Object System.Drawing.Point(620, 270)
$remainingEtaLabel.Size = New-Object System.Drawing.Size(340, 24)
$remainingEtaLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)

$logPathLabel = New-Object System.Windows.Forms.Label
$logPathLabel.Text = "Log: $logPath"
$logPathLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$logPathLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
$logPathLabel.Location = New-Object System.Drawing.Point(26, 302)
$logPathLabel.Size = New-Object System.Drawing.Size(960, 24)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(26, 332)
$progressBar.Size = New-Object System.Drawing.Size(970, 18)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.Value = 2

$metricsPanel = New-Object System.Windows.Forms.GroupBox
$metricsPanel.Text = "Live counters"
$metricsPanel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$metricsPanel.Location = New-Object System.Drawing.Point(26, 366)
$metricsPanel.Size = New-Object System.Drawing.Size(970, 184)
$metricsPanel.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)

$metricDefinitions = @(
    @{ Name = "synced_new"; Label = "New synced" },
    @{ Name = "skipped_existing"; Label = "Skipped" },
    @{ Name = "attachments_downloaded"; Label = "Attachments" },
    @{ Name = "catchup_unread_synced"; Label = "Catch-up mail" },
    @{ Name = "catchup_attachments_downloaded"; Label = "Catch-up atts" },
    @{ Name = "normalized_loaded"; Label = "Loaded" },
    @{ Name = "attachments_analyzed"; Label = "Analyzed" },
    @{ Name = "classified"; Label = "Classified" },
    @{ Name = "processed"; Label = "Processed" },
    @{ Name = "failed"; Label = "Failed" },
    @{ Name = "needs_decision"; Label = "Needs decision" }
)

$metricValueLabels = @{}
for ($i = 0; $i -lt $metricDefinitions.Count; $i++) {
    $column = $i % 4
    $row = [int][Math]::Floor($i / 4)
    $left = 18 + ($column * 232)
    $top = 28 + ($row * 42)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $metricDefinitions[$i].Label
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    $label.Location = New-Object System.Drawing.Point($left, $top)
    $label.Size = New-Object System.Drawing.Size(140, 20)

    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = "0"
    $valueLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Bold)
    $valueLabel.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)
    $valueLabel.Location = New-Object System.Drawing.Point($left, ($top + 16))
    $valueLabel.Size = New-Object System.Drawing.Size(180, 24)

$metricsPanel.Controls.AddRange(@($label, $valueLabel))
    $metricValueLabels[$metricDefinitions[$i].Name] = $valueLabel
}

$stagePanel = New-Object System.Windows.Forms.GroupBox
$stagePanel.Text = "Pipeline stages"
$stagePanel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$stagePanel.Location = New-Object System.Drawing.Point(26, 568)
$stagePanel.Size = New-Object System.Drawing.Size(970, 280)
$stagePanel.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)

$stageDefinitions = @(
    @{ Name = "sync"; Title = "1. Sync"; Description = "Read Outlook and write raw artifacts" },
    @{ Name = "catchup"; Title = "2. Catch-up"; Description = "Download historical unread mail that is not yet in raw" },
    @{ Name = "process"; Title = "3. Process"; Description = "Launch the raw mail processing pipeline" },
    @{ Name = "normalize"; Title = "4. Normalize"; Description = "Load raw mail into structured records" },
    @{ Name = "attachments"; Title = "5. Attachments"; Description = "Analyze Excel, PDF, image, and other attachments" },
    @{ Name = "classify"; Title = "6. Classify"; Description = "Assign process type, labels, and urgency" },
    @{ Name = "decision"; Title = "7. Decision"; Description = "Apply decision matrix and refresh queue" },
    @{ Name = "backlog"; Title = "8. Backlog"; Description = "Scan previous years for unread and not-in-raw mail" }
)

$defaultStageAverageSeconds = @{
    sync = 4
    catchup = 70
    process = 2
    normalize = 3
    attachments = 8
    classify = 2
    decision = 2
    backlog = 120
}

$stageTimingProfile = @{}
if (Test-Path -LiteralPath $timingPath) {
    try {
        $timingPayload = Get-Content -LiteralPath $timingPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -ne $timingPayload -and $null -ne $timingPayload.stages) {
            foreach ($property in $timingPayload.stages.PSObject.Properties) {
                $stageTimingProfile[$property.Name] = @{
                    avgSeconds = [double]$property.Value.avgSeconds
                    samples = [int]$property.Value.samples
                }
            }
        }
    }
    catch {
    }
}

foreach ($stage in $stageDefinitions) {
    if (-not $stageTimingProfile.ContainsKey($stage.Name)) {
        $defaultSeconds = 10
        if ($defaultStageAverageSeconds.ContainsKey($stage.Name)) {
            $defaultSeconds = [int]$defaultStageAverageSeconds[$stage.Name]
        }
        $stageTimingProfile[$stage.Name] = @{
            avgSeconds = [double]$defaultSeconds
            samples = 0
        }
    }
}

$stageStatusLabels = @{}
$stageDetailLabels = @{}
$stageBaseStates = @{}
$stageBaseDetails = @{}
$rowTop = 28
foreach ($stage in $stageDefinitions) {
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = $stage.Title
    $nameLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
    $nameLabel.Location = New-Object System.Drawing.Point(18, $rowTop)
    $nameLabel.Size = New-Object System.Drawing.Size(120, 22)

    $statusValueLabel = New-Object System.Windows.Forms.Label
    $statusValueLabel.Text = "pending"
    $statusValueLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $statusValueLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    $statusValueLabel.Location = New-Object System.Drawing.Point(150, $rowTop)
    $statusValueLabel.Size = New-Object System.Drawing.Size(120, 22)

    $detailValueLabel = New-Object System.Windows.Forms.Label
    $detailValueLabel.Text = $stage.Description
    $detailValueLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $detailValueLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)
    $detailValueLabel.Location = New-Object System.Drawing.Point(280, $rowTop)
    $detailValueLabel.Size = New-Object System.Drawing.Size(660, 22)

    $stagePanel.Controls.AddRange(@($nameLabel, $statusValueLabel, $detailValueLabel))
    $stageStatusLabels[$stage.Name] = $statusValueLabel
    $stageDetailLabels[$stage.Name] = $detailValueLabel
    $stageBaseStates[$stage.Name] = "pending"
    $stageBaseDetails[$stage.Name] = $stage.Description
    $rowTop += 30
}

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(26, 868)
$logBox.Size = New-Object System.Drawing.Size(970, 112)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(28, 6, 9)
$logBox.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)
$logBox.Font = New-Object System.Drawing.Font("Consolas", 10)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Stop Cycle"
$stopButton.Location = New-Object System.Drawing.Point(26, 994)
$stopButton.Size = New-Object System.Drawing.Size(150, 34)
$stopButton.BackColor = [System.Drawing.Color]::FromArgb(150, 30, 45)
$stopButton.ForeColor = [System.Drawing.Color]::White
$stopButton.FlatStyle = "Flat"

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Location = New-Object System.Drawing.Point(846, 994)
$closeButton.Size = New-Object System.Drawing.Size(150, 34)
$closeButton.BackColor = [System.Drawing.Color]::FromArgb(52, 12, 18)
$closeButton.ForeColor = [System.Drawing.Color]::White
$closeButton.FlatStyle = "Flat"
$closeButton.Enabled = $false

$form.Controls.AddRange(@(
    $title,
    $subtitle,
    $statusLabel,
    $purposeLabel,
    $elapsedLabel,
    $currentStageLabel,
    $stageActivityLabel,
    $childStateLabel,
    $pipelineStateLabel,
    $activityStateLabel,
    $lastEventLabel,
    $currentStageEtaLabel,
    $nextStageEtaLabel,
    $remainingEtaLabel,
    $logPathLabel,
    $progressBar,
    $metricsPanel,
    $stagePanel,
    $logBox,
    $stopButton,
    $closeButton
))

$startTime = Get-Date
$process = $null
$lastScreenshotAt = [datetime]::MinValue
$screenshotIndex = 0
$heartbeatFrames = @("|", "/", "-", "\")
$heartbeatIndex = 0
$lastActivityAt = Get-Date
$lastActivityDetail = "waiting to start"
$stdoutLineCount = 0
$stderrLineCount = 0
$exitObservedAt = $null
$postExitStableTicks = 0
$lastObservedOutputLineCount = 0
$runFinalized = $false
$activeStageName = ""
$activeStageStartedAt = $null
$stopRequested = $false
$lastRunSucceeded = $false
$stageStartedAts = @{}
$stageCompletedDurations = @{}
$script:lastLogWriteError = ""

function Append-SharedLogLine {
    param(
        [string]$Path,
        [string]$Line
    )

    $maxAttempts = 6
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $fileStream = $null
        $writer = $null
        try {
            $fileStream = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = New-Object System.IO.StreamWriter($fileStream, ([System.Text.UTF8Encoding]::new($false)))
            $writer.WriteLine($Line)
            $writer.Flush()
            $script:lastLogWriteError = ""
            return $true
        }
        catch [System.IO.IOException] {
            if ($attempt -ge $maxAttempts) {
                $script:lastLogWriteError = $_.Exception.Message
                return $false
            }
            Start-Sleep -Milliseconds (75 * $attempt)
        }
        catch {
            $script:lastLogWriteError = $_.Exception.Message
            return $false
        }
        finally {
            if ($writer -ne $null) {
                $writer.Dispose()
            }
            elseif ($fileStream -ne $null) {
                $fileStream.Dispose()
            }
        }
    }

    return $false
}

function Add-LogLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $timestamped = "{0}  {1}" -f (Get-Date -Format "HH:mm:ss"), $Line
    $writeOk = Append-SharedLogLine -Path $logPath -Line $timestamped
    $logBox.AppendText($timestamped + [Environment]::NewLine)
    if (-not $writeOk -and -not [string]::IsNullOrWhiteSpace($script:lastLogWriteError)) {
        $warning = "{0}  [monitor-warning] log file append skipped: {1}" -f (Get-Date -Format "HH:mm:ss"), $script:lastLogWriteError
        $logBox.AppendText($warning + [Environment]::NewLine)
    }
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
}

function Format-ShortDuration {
    param([double]$Seconds)

    if ($Seconds -lt 0) {
        return "n/a"
    }

    $rounded = [timespan]::FromSeconds([Math]::Max(0, [Math]::Round($Seconds)))
    return "{0:hh\:mm\:ss}" -f $rounded
}

function Get-StageAverageSeconds {
    param([string]$Stage)

    if ($script:stageTimingProfile.ContainsKey($Stage)) {
        return [double]$script:stageTimingProfile[$Stage].avgSeconds
    }

    if ($defaultStageAverageSeconds.ContainsKey($Stage)) {
        return [double]$defaultStageAverageSeconds[$Stage]
    }

    return 10.0
}

function Save-TimingProfile {
    $payload = @{
        updatedUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        stages = @{}
    }

    foreach ($stage in $stageDefinitions) {
        $stageName = $stage.Name
        $profile = $script:stageTimingProfile[$stageName]
        $payload.stages[$stageName] = @{
            avgSeconds = [Math]::Round([double]$profile.avgSeconds, 2)
            samples = [int]$profile.samples
        }
    }

    $json = $payload | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $timingPath -Value $json -Encoding UTF8
}

function Update-TimingProfileFromRun {
    foreach ($stage in $stageDefinitions) {
        $stageName = $stage.Name
        if (-not $script:stageCompletedDurations.ContainsKey($stageName)) {
            continue
        }

        $duration = [timespan]$script:stageCompletedDurations[$stageName]
        $seconds = [Math]::Max(1, [Math]::Round($duration.TotalSeconds, 2))
        $profile = $script:stageTimingProfile[$stageName]
        $samples = [int]$profile.samples
        $average = [double]$profile.avgSeconds

        if ($samples -le 0) {
            $profile.avgSeconds = $seconds
            $profile.samples = 1
        }
        else {
            $profile.avgSeconds = (($average * $samples) + $seconds) / ($samples + 1)
            $profile.samples = $samples + 1
        }

        $script:stageTimingProfile[$stageName] = $profile
    }

    Save-TimingProfile
}

function Register-Activity {
    param([string]$Detail)

    $script:lastActivityAt = Get-Date
    $script:lastActivityDetail = $Detail
    $lastEventLabel.Text = "Last event: {0} - {1}" -f $script:lastActivityAt.ToString("HH:mm:ss"), $Detail
}

function Get-NextPendingStageName {
    if ([string]::IsNullOrWhiteSpace($script:activeStageName)) {
        foreach ($stage in $stageDefinitions) {
            $stageName = $stage.Name
            $state = $script:stageBaseStates[$stageName]
            if ($null -ne $state -and $state.ToString().ToLowerInvariant() -eq "pending") {
                return $stageName
            }
        }
        return ""
    }

    $foundActive = $false
    foreach ($stage in $stageDefinitions) {
        $stageName = $stage.Name
        if ($stageName -eq $script:activeStageName) {
            $foundActive = $true
            continue
        }
        if (-not $foundActive) {
            continue
        }

        $state = $script:stageBaseStates[$stageName]
        if ($null -ne $state -and $state.ToString().ToLowerInvariant() -eq "pending") {
            return $stageName
        }
    }

    return ""
}

function Get-RemainingEtaSeconds {
    $remainingSeconds = 0.0

    foreach ($stage in $stageDefinitions) {
        $stageName = $stage.Name
        $state = ""
        if ($script:stageBaseStates.ContainsKey($stageName) -and $null -ne $script:stageBaseStates[$stageName]) {
            $state = $script:stageBaseStates[$stageName].ToString().ToLowerInvariant()
        }

        if ($stageName -eq $script:activeStageName -and $state -eq "running" -and $null -ne $script:activeStageStartedAt) {
            $averageSeconds = Get-StageAverageSeconds $stageName
            $elapsedSeconds = ((Get-Date) - $script:activeStageStartedAt).TotalSeconds
            $remainingSeconds += [Math]::Max(0, $averageSeconds - $elapsedSeconds)
            continue
        }

        if ($state -eq "pending") {
            $remainingSeconds += Get-StageAverageSeconds $stageName
        }
    }

    return $remainingSeconds
}

function Update-EtaVisuals {
    if ($script:runFinalized) {
        $currentStageEtaLabel.Text = "Current stage ETA: all processes stopped"
        $nextStageEtaLabel.Text = "Next stage ETA: no next stage"
        $remainingEtaLabel.Text = "Remaining ETA: 00:00:00"
        return
    }

    if (-not $script:process) {
        $currentStageEtaLabel.Text = "Current stage ETA: waiting to start"
        $nextStageEtaLabel.Text = "Next stage ETA: waiting to start"
        $remainingEtaLabel.Text = "Remaining ETA: n/a"
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($script:activeStageName) -and $null -ne $script:activeStageStartedAt) {
        $averageSeconds = Get-StageAverageSeconds $script:activeStageName
        $elapsedSeconds = ((Get-Date) - $script:activeStageStartedAt).TotalSeconds
        $currentRemaining = [Math]::Max(0, $averageSeconds - $elapsedSeconds)
        $currentStageEtaLabel.Text = "Current stage ETA: ~{0}" -f (Format-ShortDuration $currentRemaining)
    }
    else {
        $currentStageEtaLabel.Text = "Current stage ETA: n/a"
    }

    $nextStageName = Get-NextPendingStageName
    if ([string]::IsNullOrWhiteSpace($nextStageName)) {
        $nextStageEtaLabel.Text = "Next stage ETA: no pending stages"
    }
    else {
        $nextStageAverage = Get-StageAverageSeconds $nextStageName
        $nextStageEtaLabel.Text = "Next stage ETA: {0} ~{1}" -f $nextStageName.ToUpperInvariant(), (Format-ShortDuration $nextStageAverage)
    }

    $remainingSeconds = Get-RemainingEtaSeconds
    $finishTime = (Get-Date).AddSeconds($remainingSeconds)
    $remainingEtaLabel.Text = "Remaining ETA: ~{0} | finish ~{1}" -f (Format-ShortDuration $remainingSeconds), $finishTime.ToString("HH:mm:ss")
}

function Update-PipelineStateVisual {
    if ($script:runFinalized) {
        if ($script:stopRequested) {
            $pipelineStateLabel.Text = "Pipeline state: STOPPED - all processes stopped"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)
        }
        elseif ($script:lastRunSucceeded) {
            $pipelineStateLabel.Text = "Pipeline state: COMPLETED - all processes stopped"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 229, 193)
        }
        else {
            $pipelineStateLabel.Text = "Pipeline state: FAILED - all processes stopped"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 145, 145)
        }
        return
    }

    if ($script:process -and -not $script:process.HasExited) {
        $idleSeconds = [int]((Get-Date) - $script:lastActivityAt).TotalSeconds
        if ($idleSeconds -le 5) {
            $pipelineStateLabel.Text = "Pipeline state: ACTIVE"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 229, 193)
        }
        elseif ($idleSeconds -le 20) {
            $pipelineStateLabel.Text = "Pipeline state: WAITING"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        }
        else {
            $pipelineStateLabel.Text = "Pipeline state: ALIVE / QUIET"
            $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)
        }
        return
    }

    if ($script:process -and $script:process.HasExited) {
        $pipelineStateLabel.Text = "Pipeline state: FINALIZING"
        $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        return
    }

    $pipelineStateLabel.Text = "Pipeline state: STOPPED - all processes stopped"
    $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)
}

function Save-DebugScreenshot {
    param([string]$Reason)

    if (-not $form.Visible) {
        return
    }

    try {
        $script:screenshotIndex += 1
        $safeReason = ($Reason -replace '[^A-Za-z0-9_-]', '-').Trim('-')
        if ([string]::IsNullOrWhiteSpace($safeReason)) {
            $safeReason = "snapshot"
        }

        $bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
        $rectangle = New-Object System.Drawing.Rectangle 0, 0, $form.Width, $form.Height
        $form.DrawToBitmap($bitmap, $rectangle)

        $fileName = "{0:D4}_{1}_{2}.png" -f $script:screenshotIndex, $safeReason, (Get-Date -Format "HHmmss")
        $targetPath = Join-Path $screenshotDir $fileName
        $bitmap.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png)

        $bitmap.Dispose()

        $script:lastScreenshotAt = Get-Date
        Add-LogLine ("Screenshot saved: {0}" -f $targetPath)
    }
    catch {
        Add-LogLine ("Screenshot failed: {0}" -f $_.Exception.Message)
    }
}

function Set-StageState {
    param(
        [string]$Stage,
        [string]$State,
        [string]$Detail
    )

    if (-not $stageStatusLabels.ContainsKey($Stage)) {
        return
    }

    $statusControl = $stageStatusLabels[$Stage]
    $detailControl = $stageDetailLabels[$Stage]
    $script:stageBaseStates[$Stage] = $State
    $script:stageBaseDetails[$Stage] = $Detail

    $statusControl.Text = $State
    $detailControl.Text = $Detail

    switch ($State.ToLowerInvariant()) {
        "running" {
            $statusControl.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        }
        "done" {
            $statusControl.ForeColor = [System.Drawing.Color]::FromArgb(160, 229, 193)
        }
        "failed" {
            $statusControl.ForeColor = [System.Drawing.Color]::FromArgb(255, 145, 145)
        }
        default {
            $statusControl.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
        }
    }

    if ($State.ToLowerInvariant() -eq "running") {
        $script:activeStageName = $Stage
        $script:activeStageStartedAt = Get-Date
        $script:stageStartedAts[$Stage] = $script:activeStageStartedAt
        $stageActivityLabel.Text = "Stage activity: $($Stage.ToUpperInvariant()) active"
        $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
    }
    elseif ($State.ToLowerInvariant() -in @("done", "failed")) {
        $stageDuration = $null
        if ($script:stageStartedAts.ContainsKey($Stage) -and $null -ne $script:stageStartedAts[$Stage]) {
            $stageDuration = (Get-Date) - $script:stageStartedAts[$Stage]
            $script:stageCompletedDurations[$Stage] = $stageDuration
            $script:stageBaseDetails[$Stage] = "{0} | done in {1}" -f $Detail, (Format-ShortDuration $stageDuration.TotalSeconds)
            $detailControl.Text = $script:stageBaseDetails[$Stage]
        }

        if ($script:activeStageName -eq $Stage) {
            $script:activeStageName = ""
            $script:activeStageStartedAt = $null
        }

        if ($State.ToLowerInvariant() -eq "done") {
            $stageActivityLabel.Text = if ($null -ne $stageDuration) {
                "Stage activity: $($Stage.ToUpperInvariant()) completed in {0}" -f (Format-ShortDuration $stageDuration.TotalSeconds)
            } else {
                "Stage activity: $($Stage.ToUpperInvariant()) completed"
            }
            $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 229, 193)
        }
        elseif ($State.ToLowerInvariant() -eq "failed") {
            $stageActivityLabel.Text = "Stage activity: $($Stage.ToUpperInvariant()) failed"
            $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 145, 145)
        }
    }

    $currentStageLabel.Text = "Current stage: $($Stage.ToUpperInvariant()) - $Detail"
    Register-Activity ("stage {0}: {1}" -f $Stage, $State)
    Update-PipelineStateVisual
    Update-EtaVisuals
    Save-DebugScreenshot ("stage-" + $Stage + "-" + $State)
}

function Set-MetricValue {
    param(
        [string]$Name,
        [string]$Value,
        [string]$Label
    )

    if ($metricValueLabels.ContainsKey($Name)) {
        $metricValueLabels[$Name].Text = $Value
    }

    Register-Activity ("metric {0} -> {1}" -f $Name, $Value)
    if (-not [string]::IsNullOrWhiteSpace($Label)) {
        Add-LogLine ("Metric {0}: {1}" -f $Label, $Value)
    }
}

function Handle-OutputLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $normalizedLine = $Line -replace '^\[(OUT|ERR)\]\s*', ''

    $stagePattern = '\[\[STAGE\|([^|]+)\|([^|]+)\|(.*)\]\]'
    if ($normalizedLine -match $stagePattern) {
        $stageName = $Matches[1]
        $stageState = $Matches[2]
        $stageDetail = $Matches[3]
        Set-StageState -Stage $stageName -State $stageState -Detail $stageDetail
        Add-LogLine ("Stage {0}: {1} - {2}" -f $stageName, $stageState, $stageDetail)
        return
    }

    $metricPattern = '\[\[METRIC\|([^|]+)\|([^|]+)\|(.*)\]\]'
    if ($normalizedLine -match $metricPattern) {
        $metricName = $Matches[1]
        $metricValue = $Matches[2]
        $metricLabel = $Matches[3]
        Set-MetricValue -Name $metricName -Value $metricValue -Label $metricLabel
        return
    }

    Add-LogLine $Line
}

function Drain-RedirectedOutput {
    param(
        [string]$Path,
        [string]$Prefix,
        [ref]$LineCount
    )

    if (-not (Test-Path $Path)) {
        return
    }

    try {
        $lines = Get-Content -Path $Path -ErrorAction Stop
    }
    catch {
        return
    }

    if ($null -eq $lines) {
        return
    }

    if ($lines -isnot [System.Array]) {
        $lines = @($lines)
    }

    for ($i = $LineCount.Value; $i -lt $lines.Count; $i++) {
        Handle-OutputLine ("[{0}] {1}" -f $Prefix, $lines[$i])
    }

    $LineCount.Value = $lines.Count
}

function Stop-ChildProcess {
    if ($script:process -and -not $script:process.HasExited) {
        try {
            $script:process.Kill()
        }
        catch {
        }
    }
}

function Update-PipelineProgressVisual {
    $stageNames = @($stageDefinitions | ForEach-Object { $_.Name })
    $totalStages = [Math]::Max($stageNames.Count, 1)
    $completedStages = 0

    foreach ($stageName in $stageNames) {
        $stageState = ""
        if ($script:stageBaseStates.ContainsKey($stageName) -and $null -ne $script:stageBaseStates[$stageName]) {
            $stageState = $script:stageBaseStates[$stageName].ToString().ToLowerInvariant()
        }
        if ($stageState -eq "done") {
            $completedStages += 1
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:activeStageName)) {
        $stageIndex = [array]::IndexOf($stageNames, $script:activeStageName)
        if ($stageIndex -ge 0) {
            $baseProgress = [int](($stageIndex / $totalStages) * 100)
            $nextProgress = [int]((($stageIndex + 1) / $totalStages) * 100)
            $pulseWindow = [Math]::Max(3, $nextProgress - $baseProgress - 2)
            $pulseValue = $baseProgress + 1 + (($script:heartbeatIndex % 4) * [Math]::Max(1, [int]($pulseWindow / 4)))
            $progressBar.Value = [Math]::Min(99, [Math]::Max(0, $pulseValue))
            return
        }
    }

    $progressBar.Value = [Math]::Min(100, [int](($completedStages / $totalStages) * 100))
}

function Update-ActiveStageVisuals {
    foreach ($stage in $stageDefinitions) {
        $stageName = $stage.Name
        $statusControl = $stageStatusLabels[$stageName]
        $detailControl = $stageDetailLabels[$stageName]
        $baseState = $script:stageBaseStates[$stageName]
        $baseDetail = $script:stageBaseDetails[$stageName]

        if ([string]::IsNullOrWhiteSpace($baseState)) {
            continue
        }

        if ($stageName -eq $script:activeStageName -and $baseState.ToLowerInvariant() -eq "running" -and $null -ne $script:activeStageStartedAt) {
            $stageElapsed = (Get-Date) - $script:activeStageStartedAt
            $quietSeconds = [int]((Get-Date) - $script:lastActivityAt).TotalSeconds
            $averageSeconds = Get-StageAverageSeconds $stageName
            $remainingSeconds = [Math]::Max(0, $averageSeconds - $stageElapsed.TotalSeconds)
            $statusControl.Text = "{0} {1}" -f $baseState, $heartbeatFrames[$script:heartbeatIndex]
            $detailControl.Text = "{0} | running {1:hh\:mm\:ss} | quiet {2}s | eta ~{3}" -f $baseDetail, $stageElapsed, $quietSeconds, (Format-ShortDuration $remainingSeconds)
            $detailControl.ForeColor = [System.Drawing.Color]::FromArgb(246, 234, 235)
            $stageActivityLabel.Text = "Stage activity: $($stageName.ToUpperInvariant()) active for {0:hh\:mm\:ss}, quiet {1}s" -f $stageElapsed, $quietSeconds
            $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        }
        else {
            $statusControl.Text = $baseState
            if ($baseState.ToLowerInvariant() -eq "pending") {
                $detailControl.Text = "{0} | avg ~{1}" -f $baseDetail, (Format-ShortDuration (Get-StageAverageSeconds $stageName))
            }
            else {
                $detailControl.Text = $baseDetail
            }
            $detailControl.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)
        }
    }
}

function Escape-SingleQuotedValue {
    param([string]$Value)

    return ($Value -replace "'", "''")
}

function Get-ObservedExitCode {
    if (-not (Test-Path -LiteralPath $exitCodePath)) {
        return $null
    }

    try {
        $rawValue = Get-Content -LiteralPath $exitCodePath -ErrorAction Stop | Select-Object -Last 1
        if ($null -eq $rawValue) {
            return $null
        }

        $trimmed = $rawValue.ToString().Trim()
        if ($trimmed -match '^-?\d+$') {
            return [int]$trimmed
        }
    }
    catch {
        return $null
    }

    return $null
}

function Finalize-Run {
    if ($script:runFinalized) {
        return
    }

    $script:runFinalized = $true
    $resolvedExitCode = Get-ObservedExitCode
    $exitCodeText = if ($null -ne $resolvedExitCode) { $resolvedExitCode } else { "unknown" }
    $timer.Stop()
    $elapsed = (Get-Date) - $startTime
    $elapsedLabel.Text = "Elapsed: {0:hh\:mm\:ss}" -f $elapsed
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    $progressBar.Value = 100
    $script:lastRunSucceeded = ($resolvedExitCode -eq 0)
    $childStateLabel.Text = "Child process: exited ({0})" -f $exitCodeText
    $childStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    $activityStateLabel.Text = "Activity: stopped"
    $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    $stageActivityLabel.Text = "Stage activity: all processes stopped"
    $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    Register-Activity ("child exited with code {0}" -f $exitCodeText)

    if ($resolvedExitCode -eq 0) {
        Update-TimingProfileFromRun
        $statusLabel.Text = "Status: completed successfully"
        $currentStageLabel.Text = "Current stage: completed, all processes stopped"
        Add-LogLine "Mail cycle finished successfully."
    }
    elseif ($script:stopRequested) {
        $statusLabel.Text = "Status: stopped by user"
        $currentStageLabel.Text = "Current stage: stopped, all processes stopped"
        Add-LogLine "Mail cycle was stopped by user."
    }
    else {
        $statusLabel.Text = "Status: failed"
        $currentStageLabel.Text = "Current stage: failed, all processes stopped"
        Add-LogLine "Mail cycle failed with exit code $exitCodeText."
    }

    Update-PipelineStateVisual
    Update-EtaVisuals
    if ($resolvedExitCode -eq 0) {
        Save-DebugScreenshot "completed"
    }
    elseif ($script:stopRequested) {
        Save-DebugScreenshot "stopped"
    }
    else {
        Save-DebugScreenshot "failed"
    }
    $stopButton.Enabled = $false
    $closeButton.Enabled = $true
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    $elapsed = (Get-Date) - $startTime
    $elapsedLabel.Text = "Elapsed: {0:hh\:mm\:ss}" -f $elapsed
    $script:heartbeatIndex = ($script:heartbeatIndex + 1) % $heartbeatFrames.Count
    Update-ActiveStageVisuals
    Update-PipelineProgressVisual
    Update-PipelineStateVisual
    Update-EtaVisuals

    if ($process -and -not $process.HasExited) {
        $childStateLabel.Text = "Child process: alive {0}" -f $heartbeatFrames[$script:heartbeatIndex]
        $childStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        $script:exitObservedAt = $null
        $script:postExitStableTicks = 0

        $idleSeconds = [int]((Get-Date) - $script:lastActivityAt).TotalSeconds
        if ($idleSeconds -le 5) {
            $activityStateLabel.Text = "Activity: active"
            $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 229, 193)
        }
        elseif ($idleSeconds -le 20) {
            $activityStateLabel.Text = "Activity: waiting ($idleSeconds s since last event)"
            $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        }
        else {
            $activityStateLabel.Text = "Activity: still alive, waiting ($idleSeconds s since last event)"
            $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)
        }
    }
    elseif ($process -and $process.HasExited) {
        $observedExitCode = Get-ObservedExitCode
        $exitCodeText = if ($null -ne $observedExitCode) { $observedExitCode } else { "pending" }
        if ($null -eq $script:exitObservedAt) {
            $script:exitObservedAt = Get-Date
            $script:postExitStableTicks = 0
            $script:lastObservedOutputLineCount = $script:stdoutLineCount + $script:stderrLineCount
            $statusLabel.Text = "Status: finalizing output"
            $currentStageLabel.Text = "Current stage: draining final output"
            $stageActivityLabel.Text = "Stage activity: finalizing output after child exit"
            $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
            Add-LogLine ("Child process exited with code {0}; draining final output." -f $exitCodeText)
        }

        $childStateLabel.Text = "Child process: exited ({0}), draining logs" -f $exitCodeText
        $childStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(214, 190, 194)
        $activityStateLabel.Text = "Activity: draining final output"
        $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
        $progressBar.Value = [Math]::Max($progressBar.Value, 96)
    }
    else {
        $childStateLabel.Text = "Child process: not started"
        $childStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
        $activityStateLabel.Text = "Activity: waiting"
        $activityStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(171, 143, 147)
    }

    if ($process -and -not $process.HasExited) {
        if ($lastScreenshotAt -eq [datetime]::MinValue -or ((Get-Date) - $lastScreenshotAt).TotalSeconds -ge $screenshotIntervalSeconds) {
            Save-DebugScreenshot "interval"
        }
    }

    Drain-RedirectedOutput -Path $stdoutPath -Prefix "OUT" -LineCount ([ref]$script:stdoutLineCount)
    Drain-RedirectedOutput -Path $stderrPath -Prefix "ERR" -LineCount ([ref]$script:stderrLineCount)

    if ($process -and $process.HasExited) {
        Drain-RedirectedOutput -Path $stdoutPath -Prefix "OUT" -LineCount ([ref]$script:stdoutLineCount)
        Drain-RedirectedOutput -Path $stderrPath -Prefix "ERR" -LineCount ([ref]$script:stderrLineCount)
        $observedLineCount = $script:stdoutLineCount + $script:stderrLineCount
        if ($observedLineCount -eq $script:lastObservedOutputLineCount) {
            $script:postExitStableTicks += 1
        }
        else {
            $script:postExitStableTicks = 0
            $script:lastObservedOutputLineCount = $observedLineCount
        }

        $secondsSinceExit = 0
        if ($null -ne $script:exitObservedAt) {
            $secondsSinceExit = ((Get-Date) - $script:exitObservedAt).TotalSeconds
        }

        $resolvedExitCode = Get-ObservedExitCode
        if ($null -ne $resolvedExitCode -and $secondsSinceExit -ge 1 -and $script:postExitStableTicks -ge 3) {
            Finalize-Run
        }
        elseif ($secondsSinceExit -ge 5) {
            Add-LogLine "Final output drain timeout reached; finalizing monitor state."
            Finalize-Run
        }
    }
})

$stopButton.Add_Click({
    $script:stopRequested = $true
    Add-LogLine "Stop requested by user."
    $statusLabel.Text = "Status: stopping"
    $currentStageLabel.Text = "Current stage: stopping child mail process"
    $stageActivityLabel.Text = "Stage activity: stop requested by user"
    $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 210, 140)
    Register-Activity "stop requested"
    Save-DebugScreenshot "stop-requested"
    Stop-ChildProcess
})

$closeButton.Add_Click({
    $form.Close()
})

$form.Add_FormClosing({
    if ($process -and -not $process.HasExited) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The mail cycle is still running. Stop it before closing?",
            "Stop running cycle",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:stopRequested = $true
            Stop-ChildProcess
        }
        else {
            $_.Cancel = $true
        }
    }
})

$form.Add_Shown({
    Add-LogLine "Starting mail cycle."
    foreach ($removedDir in $removedScreenshotDirs) {
        Add-LogLine ("Deleted stale screenshot folder: {0}" -f $removedDir)
    }
    Add-LogLine ("Stdout: {0}" -f $stdoutPath)
    Add-LogLine ("Stderr: {0}" -f $stderrPath)
    Add-LogLine ("Exit code: {0}" -f $exitCodePath)
    $statusLabel.Text = "Status: running"
    $currentStageLabel.Text = "Current stage: launching Python mail cycle"
    $stageActivityLabel.Text = "Stage activity: launching Python mail cycle"
    $stageActivityLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
    $pipelineStateLabel.Text = "Pipeline state: STARTING"
    $pipelineStateLabel.ForeColor = [System.Drawing.Color]::FromArgb(243, 195, 202)
    $currentStageEtaLabel.Text = "Current stage ETA: starting"
    $nextStageEtaLabel.Text = "Next stage ETA: sync ~{0}" -f (Format-ShortDuration (Get-StageAverageSeconds "sync"))
    $remainingEtaLabel.Text = "Remaining ETA: ~{0}" -f (Format-ShortDuration (Get-RemainingEtaSeconds))
    Register-Activity "python cycle launch requested"
    Save-DebugScreenshot "started"

    $escapedRepoRoot = Escape-SingleQuotedValue $repoRoot
    $escapedPython = Escape-SingleQuotedValue $python
    $escapedSyncScript = Escape-SingleQuotedValue $syncScript
    $escapedExitCodePath = Escape-SingleQuotedValue $exitCodePath
    $escapedSyncArgs = Escape-SingleQuotedValue $SyncArgs
    $wrapperCommand = "& { Set-Location -LiteralPath '$escapedRepoRoot'; & '$escapedPython' -u '$escapedSyncScript' $escapedSyncArgs; `$code = `$LASTEXITCODE; Set-Content -LiteralPath '$escapedExitCodePath' -Value `$code; exit `$code }"
    $script:process = Start-Process `
        -FilePath $powershellExe `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $wrapperCommand) `
        -WorkingDirectory $repoRoot `
        -WindowStyle Hidden `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -PassThru

    $timer.Start()
})

[void]$form.ShowDialog()
