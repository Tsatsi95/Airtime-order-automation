$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "paths.ps1")
$logDir = Join-Path $MlForecastRoot "logs"

$monthPrefix = Get-Date -Format "yyyy-MM"
$thisMonthLogs = Get-ChildItem -Path $logDir -Filter "run_$monthPrefix*.log" -ErrorAction SilentlyContinue

$succeeded = $false
foreach ($log in $thisMonthLogs) {
    $content = Get-Content $log.FullName -Raw -ErrorAction SilentlyContinue
    if ($content -match "Pipeline completed successfully") {
        $succeeded = $true
        break
    }
}

if (-not $succeeded) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "The monthly forecast pipeline (ingest -> features -> train -> reorder) has not completed successfully yet this month ($monthPrefix).`n`n" +
        "This is a known issue with the scheduled automation - it may not have run at all, or died partway through (see HANDOVER.md).`n`n" +
        "Please run run_pipeline.ps1 manually (or ask Claude, if available).",
        "Stockout Alert Watchdog - Monthly pipeline has not run",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}
