@echo off
echo   Launching AJAS via PowerShell...
powershell -ExecutionPolicy Bypass -NoExit -File "%~dp0Open AJAS Windows.ps1"
if errorlevel 1 (
    echo.
    echo   PowerShell failed. Error code: %errorlevel%
    pause
)
