$ErrorActionPreference = "Stop"
$attachment = "C:\Users\TsatsiLetsoara\OneDrive - Flash\General - Airtime\Airtime Database\Templates\2023 Templates\Stock and pinless forecast\Daily Stock Forecast\ml_forecast\output\priority_order.xlsx"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show(
    "This month's stock forecast pipeline has run - time to check priority_order.xlsx for stockouts and critical SKUs.`n`n$attachment",
    "Reminder: review this month's priority order sheet",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
