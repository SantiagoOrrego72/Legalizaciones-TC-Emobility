@echo off
rem Instala PowerShell 7 (portatil si no hay permisos de administrador), PnP.PowerShell y, en el perfil Completo, Pester y openpyxl.
rem Equipo de la empresa (solo administracion):  Instalar-Prerrequisitos.cmd -Perfil Administracion
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "src\4-infraestructura\herramientas\Instalar-Prerrequisitos.ps1" %*
pause
