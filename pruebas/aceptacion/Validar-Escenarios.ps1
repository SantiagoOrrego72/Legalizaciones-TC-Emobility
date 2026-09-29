#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' }
<#
.SYNOPSIS
    Revisa en SharePoint, SIN hacer cambios, el resultado de los escenarios de prueba de aceptación.

.DESCRIPTION
    Para cada escenario de Escenarios.json busca su caso por la etiqueta "Prueba <id>" del asunto
    y comprueba lo que se puede comprobar en SharePoint:
      - que exista (o que NO exista, en el correo que no es de legalización),
      - formato del ID, Estado, Observaciones, Nombre Archivo y Fecha Cierre,
      - cantidad de archivos con ese ID en la biblioteca y la carpeta de estado donde están,
      - para el posible duplicado, que el aviso cite el ID del caso original,
      - para el volumen, que todos los casos existan y tengan ID distinto.
    Los correos recibidos por el colaborador y por Contabilidad se revisan a mano (ver PLAN-DE-PRUEBAS.md).
    Deja una tabla, un CSV en logs/ y termina con código 1 si algo no coincide.

.PARAMETER Escenarios
    Ids a validar. Por defecto, todos los que no son manuales (E01-E11 y E14).
    Después de hacer los cambios manuales de cierre: -Escenarios E12,E13

.EXAMPLE
    ./pruebas/aceptacion/Validar-Escenarios.ps1

.EXAMPLE
    ./pruebas/aceptacion/Validar-Escenarios.ps1 -Escenarios E12, E13
#>
[CmdletBinding()]
param(
    [string[]]$Escenarios,
    [string]$RutaConfig = (Join-Path $PSScriptRoot '..\..\config\Config.Legalizaciones.psd1'),
    [string]$RutaEscenarios = (Join-Path $PSScriptRoot 'Escenarios.json')
)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $raiz 'src\4-infraestructura\sharepoint\Comun.Legalizaciones.ps1')

$contexto = Import-ConfigLegalizaciones -RutaConfig $RutaConfig -RutaEsquema (Join-Path $raiz 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
$esq = $contexto.Esquema
$dominio = Get-Content (Join-Path $raiz 'src\1-dominio\dominio.json') -Raw -Encoding utf8 | ConvertFrom-Json
$catalogo = @((Get-Content $RutaEscenarios -Raw -Encoding utf8 | ConvertFrom-Json).escenarios)
if (-not $Escenarios) { $Escenarios = @($catalogo | Where-Object { -not $_.PSObject.Properties['manual'] } | ForEach-Object { $_.id }) }
$patronId = '^{0}-\d{{6}}-\d{{{1},}}$' -f [regex]::Escape($dominio.caso.prefijoId), $dominio.caso.digitosConsecutivo

$script:Resultados = New-Object System.Collections.Generic.List[object]
$script:IdsPorEtiqueta = @{}

function Add-Resultado([string]$Escenario, [string]$Verificacion, $Esperado, $Actual, [bool]$Cumple) {
    $script:Resultados.Add([pscustomobject]@{
        Escenario = $Escenario; Verificacion = $Verificacion; Esperado = [string]$Esperado; Actual = [string]$Actual
        Resultado = $(if ($Cumple) { 'OK' } else { 'FALLA' })
    })
}

function Get-CasosPorEtiqueta([string]$Etiqueta) {
    $caml = "<View><Query><Where><Contains><FieldRef Name='Asunto'/><Value Type='Text'>$Etiqueta</Value></Contains></Where><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>10</RowLimit></View>"
    return @(Get-PnPListItem -List $esq.Lista.Titulo -Query $caml)
}

function Get-ArchivosCaso([string]$Id) {
    $caml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='IdLegalizacion'/><Value Type='Text'>$Id</Value></Eq></Where></Query></View>"
    return @(Get-PnPListItem -List $esq.Biblioteca.Titulo -Query $caml)
}

# Valida el caso más reciente cuyo asunto contiene la etiqueta.
function Test-Caso([string]$Escenario, [string]$Etiqueta, $Esperado) {
    $casos = Get-CasosPorEtiqueta $Etiqueta
    if (-not $Esperado.registro) {
        Add-Resultado $Escenario "No se registró ('$Etiqueta')" 0 $casos.Count ($casos.Count -eq 0)
        return
    }
    Add-Resultado $Escenario "Caso registrado ('$Etiqueta')" 'Sí' $(if ($casos.Count) { 'Sí' } else { 'No' }) ($casos.Count -gt 0)
    if ($casos.Count -eq 0) { return }

    $valores = $casos[0].FieldValues
    $id = [string]$valores['Title']
    $script:IdsPorEtiqueta[$Etiqueta] = $id
    Add-Resultado $Escenario 'Formato del ID' $patronId $id ($id -match $patronId)
    if ($Esperado.PSObject.Properties['estado']) {
        Add-Resultado $Escenario 'Estado' $Esperado.estado $valores['Estado'] ([string]$valores['Estado'] -eq $Esperado.estado)
    }
    $observaciones = [string]$valores['Observaciones']
    if ($Esperado.PSObject.Properties['observacionesVacias'] -and $Esperado.observacionesVacias) {
        Add-Resultado $Escenario 'Observaciones vacías' '(vacío)' $observaciones ([string]::IsNullOrWhiteSpace($observaciones))
    }
    foreach ($fragmento in @($Esperado.observacionesContienen)) {
        if (-not $fragmento) { continue }
        Add-Resultado $Escenario "Observaciones contienen '$fragmento'" 'Sí' $observaciones ($observaciones.Contains($fragmento))
    }
    if ($Esperado.PSObject.Properties['duplicadoDe']) {
        $original = $script:IdsPorEtiqueta["Prueba $($Esperado.duplicadoDe)"]
        if (-not $original) { $original = [string](Get-CasosPorEtiqueta "Prueba $($Esperado.duplicadoDe)" | Select-Object -First 1).FieldValues['Title'] }
        Add-Resultado $Escenario "El aviso cita el caso de $($Esperado.duplicadoDe)" $original $observaciones ([bool]$original -and $observaciones.Contains($original))
    }
    if ($Esperado.PSObject.Properties['fechaCierre'] -and $Esperado.fechaCierre) {
        Add-Resultado $Escenario 'Fecha Cierre registrada' 'Sí' $valores['FechaCierre'] ($null -ne $valores['FechaCierre'])
    }
    if ($Esperado.PSObject.Properties['archivosGuardados']) {
        $nombres = @(([string]$valores['NombreArchivo']).Replace("`r", '').Split("`n") | Where-Object { $_ })
        Add-Resultado $Escenario 'Nombre Archivo (cantidad)' $Esperado.archivosGuardados $nombres.Count ($nombres.Count -eq [int]$Esperado.archivosGuardados)
        $archivos = Get-ArchivosCaso $id
        Add-Resultado $Escenario 'Archivos con este ID en la biblioteca' $Esperado.archivosGuardados $archivos.Count ($archivos.Count -eq [int]$Esperado.archivosGuardados)
        if ($Esperado.PSObject.Properties['carpeta'] -and $archivos.Count -gt 0) {
            $fuera = @($archivos | Where-Object { ([string]$_.FieldValues['FileRef']) -notmatch ('/{0}/{1}/' -f [regex]::Escape($esq.Biblioteca.Url), [regex]::Escape($Esperado.carpeta)) })
            Add-Resultado $Escenario "Archivos en la carpeta $($Esperado.carpeta)" 0 $fuera.Count ($fuera.Count -eq 0)
        }
    }
}

$rutaLog = Start-RegistroEjecucion -Contexto $contexto -Prefijo 'validacion-escenarios'
$fallas = 0
try {
    Connect-SitioLegalizaciones -Contexto $contexto
    foreach ($idEscenario in $Escenarios) {
        $escenario = $catalogo | Where-Object { $_.id -eq $idEscenario } | Select-Object -First 1
        if (-not $escenario) { Write-Resultado 'AVISO' "El escenario $idEscenario no está en Escenarios.json"; continue }
        Write-Paso "$($escenario.id) · $($escenario.nombre)"

        if ($escenario.PSObject.Properties['volumen']) {
            $ids = @()
            for ($n = 1; $n -le [int]$escenario.volumen.cantidad; $n++) {
                $etiqueta = 'Prueba V{0:D2}' -f $n
                Test-Caso $escenario.id $etiqueta $escenario.esperado
                if ($script:IdsPorEtiqueta.ContainsKey($etiqueta)) { $ids += $script:IdsPorEtiqueta[$etiqueta] }
            }
            $distintos = @($ids | Select-Object -Unique).Count
            Add-Resultado $escenario.id 'IDs distintos en el volumen' $escenario.volumen.cantidad $distintos ($distintos -eq [int]$escenario.volumen.cantidad)
        }
        elseif ($escenario.PSObject.Properties['casoDe']) {
            Test-Caso $escenario.id "Prueba $($escenario.casoDe)" $escenario.esperado
        }
        else {
            Test-Caso $escenario.id "Prueba $($escenario.id)" $escenario.esperado
        }
    }

    Write-Paso 'Resultado'
    $script:Resultados | Format-Table Escenario, Verificacion, Esperado, Actual, Resultado -AutoSize -Wrap | Out-Host
    $rutaCsv = [System.IO.Path]::ChangeExtension($rutaLog, '.csv')
    $script:Resultados | Export-Csv -LiteralPath $rutaCsv -NoTypeInformation -Encoding utf8BOM
    $fallas = @($script:Resultados | Where-Object { $_.Resultado -eq 'FALLA' }).Count
    Write-Host ('  OK: {0}   FALLA: {1}' -f @($script:Resultados | Where-Object { $_.Resultado -eq 'OK' }).Count, $fallas)
    Write-Host "  Resultados en CSV: $rutaCsv"
    Write-Host '  Revise además a mano los correos recibidos (columna "Correo" de PLAN-DE-PRUEBAS.md).'
}
catch {
    Write-Resultado 'ERROR' $_.Exception.Message
    throw
}
finally {
    Stop-Transcript | Out-Null
}
if ($fallas -gt 0) { exit 1 }
