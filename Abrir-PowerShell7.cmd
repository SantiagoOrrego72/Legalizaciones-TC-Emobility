@echo off
rem Abre PowerShell 7 (instalado o portatil) en la carpeta del proyecto.
setlocal
set "PWSH="
where pwsh >nul 2>nul && set "PWSH=pwsh"
if not defined PWSH if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PWSH if exist "%LOCALAPPDATA%\Programs\PowerShell-7\pwsh.exe" set "PWSH=%LOCALAPPDATA%\Programs\PowerShell-7\pwsh.exe"
if not defined PWSH goto :sinpwsh
cd /d "%~dp0"
"%PWSH%" -NoExit -NoLogo -Command "Write-Host ('Legalizaciones TC - PowerShell ' + $PSVersionTable.PSVersion) -ForegroundColor Cyan; Write-Host 'Guia: README.md' -ForegroundColor Cyan"
goto :eof

:sinpwsh
echo No se encontro PowerShell 7. Ejecute primero Instalar-Prerrequisitos.cmd
pause
