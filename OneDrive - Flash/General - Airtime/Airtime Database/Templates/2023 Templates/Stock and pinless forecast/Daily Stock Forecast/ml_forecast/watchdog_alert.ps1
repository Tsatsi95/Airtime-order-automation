param(
    [Parameter(Mandatory=$true)][string]$Label  # "Morning" or "Afternoon"
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "paths.ps1")
$logDir = Join-Path $MlForecastRoot "logs"

$today = Get-Date -Format "yyyy-MM-dd"
$todaysLogs = Get-ChildItem -Path $logDir -Filter "alert_$today*.log" -ErrorAction SilentlyContinue

$succeeded = $false
foreach ($log in $todaysLogs) {
    $content = Get-Content $log.FullName -Raw -ErrorAction SilentlyContinue
    if ($content -match "Email sent to:" -or $content -match "skipping email") {
        $succeeded = $true
        break
    }
}

if (-not $succeeded) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "The $Label stockout alert did not complete successfully today ($today).`n`n" +
        "Either it never ran, or it failed partway through (a known issue with the Excel/Outlook automation - see HANDOVER.md).`n`n" +
        "Please run team_stockout_alert.ps1 manually (or ask Claude, if available) to make sure the team gets today's numbers.",
        "Stockout Alert Watchdog - $Label alert did not send",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}
