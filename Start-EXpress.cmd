@echo off
REM Quickstart launcher for EXpress
REM Opens an elevated PowerShell that stays open after the script exits or fails.
powershell.exe -NoExit -ExecutionPolicy Bypass -NoProfile -File "%~dp0EXpress.ps1" %*
