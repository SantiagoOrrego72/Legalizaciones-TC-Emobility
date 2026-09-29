@echo off
rem Ejecuta todas las pruebas automaticas (no necesitan tenant). Use: Ejecutar-Pruebas.cmd -SinExcel para omitir las de Excel.
setlocal
set "PWSH="
where pwsh >nul 2>nul && set "PWSH=pwsh"
if not defined PWSH if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not defined PWSH if exist "%LOCALAPPDATA%\Programs\PowerShell-7\pwsh.exe" set "PWSH=%LOCALAPPDATA%\Programs\PowerShell-7\pwsh.exe"
if not defined PWSH goto :sinpwsh
cd /d "%~dp0"
"%PWSH%" -NoProfile -File "pruebas\Invoke-Pruebas.ps1" %*
pause
goto :eof

:sinpwsh
echo No se encontro PowerShell 7. Ejecute primero Instalar-Prerrequisitos.cmd
pause
