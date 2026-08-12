$ErrorActionPreference = "Stop"

$python = "C:\Users\TsatsiLetsoara\AppData\Local\Programs\Python\Python312\python.exe"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "src"
$logDir = Join-Path $root "logs"

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("run_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))

# Windows PowerShell 5.1's Tee-Object has no -Encoding parameter (that's a
# PowerShell 6+ addition) — Out-File does support it in 5.1, so use that instead.
function Write-Log($line) {
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

function Send-FailureAlert($subject, $body) {
    # Outlook COM automation (even wrapped with a background-job timeout) has
    # proven unreliable when invoked from a Task-Scheduler-triggered process in
    # this environment — observed hangs, and once a full system stall requiring
    # a restart. Dropped entirely in favor of the one mechanism independently
    # confirmed to work reliably under Task Scheduler here: a local popup.
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        $body, $subject,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

Set-Location $src

$steps = @("ingest.py", "features.py", "train.py", "reorder.py", "priority_order.py")

# scheduled trigger is the 1st of the month at 06:00 — if this run starts later
# than that (StartWhenAvailable catching up a missed run, or a manual re-run),
# flag it so a delay isn't mistaken for "ran on time and looks stale for no reason"
$now = Get-Date
$expectedStart = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 6 -Minute 0 -Second 0
$LATE_RUN_GRACE_HOURS = 1
$lateBy = $now - $expectedStart
if ($lateBy.TotalHours -gt $LATE_RUN_GRACE_HOURS) {
    Write-Log "*** LATE RUN: started $([math]::Round($lateBy.TotalHours, 1)) hours after the expected 06:00 on the 1st (expected $expectedStart, actual $now) ***"
}

Write-Log "=== Monthly forecast pipeline run: $(Get-Date -Format o) ==="

foreach ($step in $steps) {
    Write-Log "--- Running $step ---"
    & $python $step 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
    if ($LASTEXITCODE -ne 0) {
        Write-Log "!!! $step failed with exit code $LASTEXITCODE - stopping pipeline !!!"

        $logTail = (Get-Content $logFile -Tail 40) -join "`r`n"
        Send-FailureAlert `
            -subject "Airtime stock forecast pipeline FAILED ($step)" `
            -body "The monthly forecast pipeline failed at step '$step' on $(Get-Date -Format o).`r`n`r`nLog file: $logFile`r`n`r`nLast 40 lines:`r`n$logTail"

        exit $LASTEXITCODE
    }
}

Write-Log "=== Pipeline completed successfully: $(Get-Date -Format o) ==="
