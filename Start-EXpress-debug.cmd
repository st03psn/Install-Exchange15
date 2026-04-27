@echo off
REM Debug-Launcher fuer EXpress
REM Startet EXpress mit deploy-example_contoso.psd1 und aktiviertem -Debug.
REM Optionale Parameter direkt anhaengen, z.B.: Start-EXpress-debug.cmd -Phase 5
powershell.exe -NoExit -ExecutionPolicy Bypass -NoProfile -File "%~dp0EXpress.ps1" -ConfigFile "%~dp0deploy-example_contoso.psd1" -Debug %*
