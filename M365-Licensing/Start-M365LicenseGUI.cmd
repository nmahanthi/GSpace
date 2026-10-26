@echo off
cd /d "%~dp0"
powershell -STA -NoProfile -ExecutionPolicy Bypass -File ".\Start-M365LicenseGUI.ps1"
