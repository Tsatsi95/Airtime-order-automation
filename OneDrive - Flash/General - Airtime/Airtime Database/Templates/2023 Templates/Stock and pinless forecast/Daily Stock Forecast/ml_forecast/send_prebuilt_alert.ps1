$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "paths.ps1")
$root = $MlForecastRoot
$recipients = @("lutho.posiliti@flash.co.za", "Khanyesa.Khota@flash.co.za", "Nontuthuko.madondo@flash.co.za")

$tableHtml = Get-Content (Join-Path $root "output\_alert_table.html") -Raw -Encoding UTF8
$subject = "Stock Alert: 9 stocked out, 0 critical"
$body = "<p>Daily stock check - please keep an eye on these:</p>$tableHtml<p>Full list: priority_order.xlsx</p>"

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = ($recipients -join "; ")
$mail.Subject = $subject
$mail.HTMLBody = $body
$mail.Send()
Write-Output "Email sent to: $($recipients -join ', ')"
