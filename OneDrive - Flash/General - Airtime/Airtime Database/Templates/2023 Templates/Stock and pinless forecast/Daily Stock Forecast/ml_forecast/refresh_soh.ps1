$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "paths.ps1")
$path = $VvsForecastPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow - suppress security/trust prompts that DisplayAlerts alone doesn't cover

try {
    $wb = $excel.Workbooks.Open($path)
    Write-Output "Opened workbook, starting RefreshAll..."

    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    $wb.Save()
    Write-Output "Refresh complete and saved."
    $wb.Close($true)
} finally {
    try {
        $excel.Quit()
    } catch {
        # COM object doesn't expose its own process ID cleanly - if Quit() is
        # rejected (Excel busy on a dialog), the caller needs to find and kill
        # the orphaned EXCEL.EXE process externally (e.g. by MainWindowTitle).
        Write-Output "Quit() failed: $_"
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
