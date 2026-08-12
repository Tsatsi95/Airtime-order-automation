$ErrorActionPreference = "Stop"
$path = "C:\Users\TsatsiLetsoara\OneDrive - Flash\General - Airtime\Airtime Database\Templates\2023 Templates\Stock and pinless forecast\Daily Stock Forecast\VVS_FORECAST.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($path)
    Write-Output "Opened workbook, starting RefreshAll..."

    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    $wb.Save()
    Write-Output "Refresh complete and saved."
    $wb.Close($true)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
