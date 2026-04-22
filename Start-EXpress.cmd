@echo off
REM Quickstart launcher for EXpress (Install-Exchange15.ps1)
REM Opens an elevated PowerShell that stays open after the script exits or fails.
powershell.exe -NoExit -ExecutionPolicy Bypass -NoProfile -Command "Set-Location 'C:\install'; .\Install-Exchange15.ps1 -Debug %*"
