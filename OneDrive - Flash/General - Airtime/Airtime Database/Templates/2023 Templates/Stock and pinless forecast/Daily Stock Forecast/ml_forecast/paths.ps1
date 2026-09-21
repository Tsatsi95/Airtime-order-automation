# Shared path/tool resolution for the ml_forecast automation scripts.
# Dot-source this at the top of every script: . (Join-Path $PSScriptRoot "paths.ps1")
# Avoids hardcoding a specific Windows username or Python install location so
# this works on a new machine. See HANDOVER.md before migrating.

function Get-OneDriveFlashRoot {
    if ($env:ONEDRIVE_FLASH_ROOT -and (Test-Path $env:ONEDRIVE_FLASH_ROOT)) {
        return $env:ONEDRIVE_FLASH_ROOT
    }
    $exact = Join-Path $env:USERPROFILE "OneDrive - Flash"
    if (Test-Path $exact) { return $exact }

    $found = Get-ChildItem -Path $env:USERPROFILE -Directory -Filter "OneDrive*Flash*" -ErrorAction SilentlyContinue
    if ($found) { return $found[0].FullName }

    throw "Could not find the 'OneDrive - Flash' folder under $env:USERPROFILE. Make sure OneDrive is installed and synced, or set the ONEDRIVE_FLASH_ROOT environment variable."
}

function Get-PythonExe {
    if ($env:PYTHON_EXE -and (Test-Path $env:PYTHON_EXE)) { return $env:PYTHON_EXE }
    $onPath = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    throw "Python was not found on PATH. Install Python 3.12+ and the packages in requirements.txt, or set the PYTHON_EXE environment variable to its full path."
}

$OneDriveFlashRoot = Get-OneDriveFlashRoot
$AirtimeDbTemplates = Join-Path $OneDriveFlashRoot "General - Airtime\Airtime Database\Templates"
$DailyStockForecastDir = Join-Path $AirtimeDbTemplates "2023 Templates\Stock and pinless forecast\Daily Stock Forecast"
$MlForecastRoot = Join-Path $DailyStockForecastDir "ml_forecast"
$MlForecastSrc = Join-Path $MlForecastRoot "src"
$VvsForecastPath = Join-Path $DailyStockForecastDir "VVS_FORECAST.xlsx"
$PythonExe = Get-PythonExe
