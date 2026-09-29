#Requires -Version 7.4
<#
.SYNOPSIS
    Ejecuta todas las pruebas automáticas del proyecto (no necesitan tenant ni conexión).

.DESCRIPTION
    - Pester (pruebas/unitarias/*.Tests.ps1):
        Dominio          reglas de negocio (ejemplos verificados)
        Arquitectura     regla de dependencias y coherencia dominio <-> adaptadores
        Aprovisionamiento scripts de SharePoint contra un SharePoint simulado
        Infraestructura  comandos y parámetros PnP contra el módulo instalado
        Excel            Power Query ejecutado de verdad con Excel de escritorio (lento; -SinExcel lo omite)
    - Python (pruebas/unitarias/test_*.py): plantilla del comprobante, datos de prueba y libro de análisis.

    Termina con código 1 si alguna prueba falla.

.EXAMPLE
    ./pruebas/Invoke-Pruebas.ps1

.EXAMPLE
    ./pruebas/Invoke-Pruebas.ps1 -SinExcel
#>
[CmdletBinding()]
param(
    # Omite las pruebas que abren Excel (tardan cerca de un minuto).
    [switch]$SinExcel
)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$unitarias = Join-Path $raiz 'pruebas\unitarias'

Import-Module Pester -MinimumVersion 5.5.0 -ErrorAction Stop
$configuracion = New-PesterConfiguration
$configuracion.Run.Path = $unitarias
$configuracion.Run.PassThru = $true
$configuracion.Output.Verbosity = 'Normal'
if ($SinExcel) { $configuracion.Run.ExcludePath = @(Join-Path $unitarias 'Excel.Tests.ps1') }

Write-Host '== Pester' -ForegroundColor Cyan
$pester = Invoke-Pester -Configuration $configuracion

Write-Host '== Python (unittest)' -ForegroundColor Cyan
$python = Get-Command python -ErrorAction SilentlyContinue
$pythonOk = $true
if (-not $python) {
    Write-Warning 'No se encontró Python: se omiten las pruebas de Python.'
}
else {
    Push-Location $raiz
    try {
        & $python.Source -m unittest discover -s 'pruebas/unitarias' -p 'test_*.py'
        $pythonOk = ($LASTEXITCODE -eq 0)
    }
    finally { Pop-Location }
}

Write-Host ''
Write-Host ('Pester: {0} pasan, {1} fallan, {2} omitidas' -f $pester.PassedCount, $pester.FailedCount, $pester.SkippedCount)
Write-Host ('Python: {0}' -f $(if (-not $python) { 'omitidas' } elseif ($pythonOk) { 'todas pasan' } else { 'HAY FALLAS' }))
if ($pester.FailedCount -gt 0 -or -not $pythonOk) {
    Write-Host 'RESULTADO: hay pruebas que fallan.' -ForegroundColor Red
    exit 1
}
Write-Host 'RESULTADO: todas las pruebas pasan.' -ForegroundColor Green
