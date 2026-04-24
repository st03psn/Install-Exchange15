@echo off
REM Quickstart launcher for EXpress
REM Opens an elevated PowerShell that stays open after the script exits or fails.
powershell.exe -NoExit -ExecutionPolicy Bypass -NoProfile -Command "Set-Location 'C:\install'; .\EXpress.ps1 %*"
