<#
.SYNOPSIS
    Envía los correos de los escenarios de prueba con Outlook de escritorio (clásico).

.DESCRIPTION
    Arma cada correo con el asunto y los adjuntos de Escenarios.json y lo envía desde la cuenta
    de Outlook abierta en este equipo (usted es el remitente) al buzón de legalizaciones.
    Pide confirmación antes de enviar. Los escenarios manuales (E12, E13) no se envían.

    Requisitos:
      - Outlook clásico instalado y con su cuenta configurada. El nuevo Outlook no permite
        automatización; en ese caso envíe los correos a mano (PLAN-DE-PRUEBAS.md, sección 4).
      - Los adjuntos de prueba generados:  python pruebas/aceptacion/generar_datos_prueba.py

    Funciona en Windows PowerShell 5.1 y PowerShell 7.

.PARAMETER Escenarios
    Ids a enviar, en orden. Por defecto E01 a E11 y E14. E10 (duplicado) debe ir después de E01.

.PARAMETER Para
    Destinatario. Por defecto, Flujos.BuzonLegalizaciones de la configuración.

.PARAMETER PausaSegundos
    Espera entre correos (para que lleguen en orden). Por defecto 5.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\pruebas\aceptacion\Enviar-CorreosPrueba.ps1

.EXAMPLE
    .\pruebas\aceptacion\Enviar-CorreosPrueba.ps1 -Escenarios E01, E10
#>
[CmdletBinding()]
param(
    [string[]]$Escenarios = @('E01', 'E02', 'E03', 'E04', 'E05', 'E06', 'E07', 'E08', 'E09', 'E10', 'E14', 'E11'),
    [string]$Para,
    [ValidateRange(0, 120)][int]$PausaSegundos = 5,
    [switch]$SinConfirmar,
    [string]$RutaConfig,
    [string]$RutaEscenarios,
    [string]$CarpetaDatos
)

$ErrorActionPreference = 'Stop'
$aqui = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $RutaConfig) { $RutaConfig = Join-Path $aqui '..\..\config\Config.Legalizaciones.psd1' }
if (-not $RutaEscenarios) { $RutaEscenarios = Join-Path $aqui 'Escenarios.json' }
if (-not $CarpetaDatos) { $CarpetaDatos = Join-Path $aqui '..\datos' }

if (-not $Para) {
    $cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig
    $Para = [string]$cfg.Flujos.BuzonLegalizaciones
}
if (-not $Para -or $Para -match '<[^>]*>') {
    throw "Indique el buzón de legalizaciones con -Para o complételo en Flujos.BuzonLegalizaciones de '$RutaConfig'."
}

$catalogo = @(([System.IO.File]::ReadAllText($RutaEscenarios, [System.Text.Encoding]::UTF8) | ConvertFrom-Json).escenarios)

# Correos a enviar, en orden
$correos = @()
foreach ($id in $Escenarios) {
    $escenario = $catalogo | Where-Object { $_.id -eq $id } | Select-Object -First 1
    if (-not $escenario) { throw "El escenario $id no está en Escenarios.json." }
    if ($escenario.PSObject.Properties['manual']) { Write-Warning "$id es manual y no se envía: $($escenario.manual)"; continue }
    if ($escenario.PSObject.Properties['volumen']) {
        for ($n = 1; $n -le [int]$escenario.volumen.cantidad; $n++) {
            $nn = '{0:D2}' -f $n
            $correos += [pscustomobject]@{
                Escenario = $id
                Asunto    = $escenario.volumen.asunto.Replace('{nn}', $nn)
                Adjuntos  = @($escenario.volumen.adjuntos | ForEach-Object { $_.Replace('{nn}', $nn) })
            }
        }
    }
    else {
        $correos += [pscustomobject]@{ Escenario = $id; Asunto = $escenario.asunto; Adjuntos = @($escenario.adjuntos) }
    }
}

# Los adjuntos deben existir
$faltantes = @($correos | ForEach-Object { $_.Adjuntos } | Select-Object -Unique | Where-Object { -not (Test-Path (Join-Path $CarpetaDatos $_)) })
if ($faltantes.Count -gt 0) {
    throw ("Faltan adjuntos en '{0}': {1}. Genérelos con: python pruebas/aceptacion/generar_datos_prueba.py" -f $CarpetaDatos, ($faltantes -join ', '))
}

Write-Host ("Se enviarán {0} correos a {1}:" -f $correos.Count, $Para) -ForegroundColor Cyan
$correos | Format-Table Escenario, Asunto, @{ Label = 'Adjuntos'; Expression = { $_.Adjuntos -join ', ' } } -AutoSize | Out-Host
if (-not $SinConfirmar) {
    $respuesta = Read-Host '¿Enviar? (S/N)'
    if ($respuesta -notmatch '^[sS]') { Write-Host 'Cancelado; no se envió nada.'; return }
}

try { $outlook = New-Object -ComObject Outlook.Application }
catch { throw 'No se pudo abrir Outlook clásico por automatización. Envíe los correos a mano según PLAN-DE-PRUEBAS.md.' }

$enviados = 0
foreach ($correo in $correos) {
    $mensaje = $outlook.CreateItem(0)   # olMailItem
    $mensaje.To = $Para
    $mensaje.Subject = $correo.Asunto
    $mensaje.Body = "Correo de prueba del escenario $($correo.Escenario) (proceso de legalización de gastos). Generado por Enviar-CorreosPrueba.ps1."
    foreach ($adjunto in $correo.Adjuntos) {
        [void]$mensaje.Attachments.Add((Resolve-Path (Join-Path $CarpetaDatos $adjunto)).Path)
    }
    $mensaje.Send()
    $enviados++
    Write-Host ("  [{0}/{1}] {2}  {3}" -f $enviados, $correos.Count, $correo.Escenario, $correo.Asunto)
    if ($PausaSegundos -gt 0 -and $enviados -lt $correos.Count) { Start-Sleep -Seconds $PausaSegundos }
}
Write-Host "Listo: $enviados correos enviados. Espere unos minutos y ejecute Validar-Escenarios.ps1." -ForegroundColor Green
