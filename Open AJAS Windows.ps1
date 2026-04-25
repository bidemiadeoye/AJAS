# AJAS Launcher - PowerShell version
$ErrorActionPreference = "Continue"
$DIR = Split-Path -Parent $MyInvocation.MyCommand.Path
$LAUNCHER = Join-Path $DIR "app\ajas_launcher.py"

Write-Host ""
Write-Host "  AJAS - Starting up" -ForegroundColor Cyan
Write-Host "  Script dir: $DIR"
Write-Host "  Launcher: $LAUNCHER"
Write-Host ""

if (-not (Test-Path $LAUNCHER)) {
    Write-Host "  ERROR: Cannot find app\ajas_launcher.py" -ForegroundColor Red
    Write-Host "  Contents of DIR:"
    Get-ChildItem $DIR | Select-Object Name
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "  Launcher found OK"

# Find Python
$python = $null
$tried = @()

$candidates = @(
    "py",
    "python",
    "python3",
    "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python39\python.exe"
)

foreach ($candidate in $candidates) {
    $tried += $candidate
    try {
        $result = & $candidate --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  Found Python: $candidate -> $result" -ForegroundColor Green
            $python = $candidate
            break
        }
    } catch {}
}

if (-not $python) {
    Write-Host "  Python not found. Tried:" -ForegroundColor Yellow
    $tried | ForEach-Object { Write-Host "    $_" }
    Write-Host ""
    Write-Host "  Installing Python 3.12..." -ForegroundColor Yellow
    $url = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe"
    $installer = "$env:TEMP\python_install.exe"
    Write-Host "  Downloading..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    (New-Object Net.WebClient).DownloadFile($url, $installer)
    Write-Host "  Installing (this takes 1-2 minutes)..."
    Start-Process $installer -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1" -Wait
    Remove-Item $installer -Force -ErrorAction SilentlyContinue
    $python = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
    Write-Host "  Python installed. Continuing..." -ForegroundColor Green
}

Write-Host ""
Write-Host "  Installing/checking libraries..."
& $python -m pip install flask openpyxl python-docx requests --quiet 2>$null
Write-Host "  Libraries ready."
Write-Host ""
Write-Host "  ======================================" -ForegroundColor Cyan
Write-Host "  AJAS is running." -ForegroundColor Green
Write-Host "  Browser will open automatically." -ForegroundColor Green
Write-Host "  Keep this window open." -ForegroundColor Green
Write-Host "  ======================================" -ForegroundColor Cyan
Write-Host ""

& $python $LAUNCHER

Write-Host ""
Write-Host "  AJAS has stopped."
Read-Host "Press Enter to close"
