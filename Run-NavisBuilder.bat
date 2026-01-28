@echo off
:: Launches the Navisworks Builder PowerShell script
:: Checks current directory and ./src/ directory

if exist "%~dp0Build-NavisModels.ps1" (
    set "SCRIPT_PATH=%~dp0Build-NavisModels.ps1"
) else if exist "%~dp0src\Build-NavisModels.ps1" (
    set "SCRIPT_PATH=%~dp0src\Build-NavisModels.ps1"
) else (
    echo [ERROR] Build-NavisModels.ps1 not found in root or src folder!
    pause
    exit /b
)

echo Starting Navisworks Automation...
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"