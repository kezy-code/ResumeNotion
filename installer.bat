@echo off
:: Change le répertoire courant vers celui du .bat
cd /d "%~dp0"

:: Lance le script PowerShell dans ce répertoire
powershell.exe -ExecutionPolicy Bypass -File "installer.ps1"

pause
