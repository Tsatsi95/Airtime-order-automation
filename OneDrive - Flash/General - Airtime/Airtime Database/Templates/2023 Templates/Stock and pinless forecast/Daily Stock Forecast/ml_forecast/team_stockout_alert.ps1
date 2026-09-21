param(
    [switch]$TestMode
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $root "paths.ps1")
$src = $MlForecastSrc
$logDir = Join-Path $root "logs"
$python = $PythonExe

# Update this list if the recipient team changes, or when handing over to a
# new owner (see HANDOVER.md).
$recipients = @("lutho.posiliti@flash.co.za", "Khanyesa.Khota@flash.co.za", "Nontuthuko.madondo@flash.co.za")
if ($TestMode) {
    $testRecipient = if ($env:TEST_ALERT_RECIPIENT) { $env:TEST_ALERT_RECIPIENT } else { "tsatsi.letsoara@flash.co.za" }
    $recipients = @($testRecipient)
}

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("alert_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))
function Write-Log($line) { $line | Out-File -FilePath $logFile -Append -Encoding utf8 }

function Send-Alert($subject, $bodyHtml) {
    Add-Type -AssemblyName System.Windows.Forms
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = ($recipients -join "; ")
        $mail.Subject = $subject
        $mail.HTMLBody = $bodyHtml
        $mail.Send()
        Write-Log "Email sent to: $($recipients -join ', ')"
    } catch {
        Write-Log "Email send FAILED: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Team stockout alert email failed to send. Check log: $logFile`n`n$_",
            "Stockout alert - SEND FAILED",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

Set-Location $src
Write-Log "=== Team stockout alert run: $(Get-Date -Format o) (TestMode=$TestMode) ==="

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "refresh_soh.ps1") 2>&1 |
    Out-File -FilePath $logFile -Append -Encoding utf8

& $python "reorder.py" 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
if ($LASTEXITCODE -ne 0) { Write-Log "reorder.py failed, aborting"; exit 1 }
& $python "priority_order.py" 2>&1 | Out-File -FilePath $logFile -Append -Encoding utf8
if ($LASTEXITCODE -ne 0) { Write-Log "priority_order.py failed, aborting"; exit 1 }

$emailScriptOutput = & $python "team_stockout_email.py" 2>&1
$emailScriptOutput | Out-File -FilePath $logFile -Append -Encoding utf8

$nStockedOut = 0
$nCritical = 0
$tableWritten = $false
foreach ($line in $emailScriptOutput) {
    if ($line -match "STOCKED_OUT=(\d+)") { $nStockedOut = [int]$Matches[1] }
    if ($line -match "CRITICAL=(\d+)") { $nCritical = [int]$Matches[1] }
    if ($line -match "TABLE_WRITTEN=1") { $tableWritten = $true }
}

if (-not $tableWritten) {
    Write-Log "No stocked-out or critical SKUs - skipping email."
    exit 0
}

$tableHtml = Get-Content (Join-Path $root "output\_alert_table.html") -Raw -Encoding UTF8
$subject = "Stock Alert: $nStockedOut stocked out, $nCritical critical" + $(if ($TestMode) { " [TEST]" } else { "" })
$body = "<p>Daily stock check - please keep an eye on these:</p>$tableHtml<p>Full list: priority_order.xlsx</p>"

Send-Alert -subject $subject -bodyHtml $body
Write-Log "=== Done: $(Get-Date -Format o) ==="
