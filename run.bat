@echo off
setlocal
REM Get the folder where this batch file is located
set "scriptdir=%~dp0"
REM Build the full path to your PowerShell script in the same folder
set "script=%scriptdir%Auto-USB-Sleep-Override.ps1"
REM Run PowerShell as admin on the script
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%script%\"' -Verb RunAs"
endlocal
