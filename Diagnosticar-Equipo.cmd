@echo off
rem Revisa si este equipo tiene lo necesario para instalar y administrar la solucion (no cambia nada).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "src\4-infraestructura\herramientas\Diagnosticar-Equipo.ps1"
pause
