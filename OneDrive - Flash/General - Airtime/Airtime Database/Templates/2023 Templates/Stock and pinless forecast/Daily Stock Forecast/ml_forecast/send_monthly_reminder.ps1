$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "paths.ps1")
$attachment = Join-Path $MlForecastRoot "output\priority_order.xlsx"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show(
    "This month's stock forecast pipeline has run - time to check priority_order.xlsx for stockouts and critical SKUs.`n`n$attachment",
    "Reminder: review this month's priority order sheet",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
