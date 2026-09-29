<#
.SYNOPSIS
    Genera docs/recorrido/recorrido.html: página autocontenida con el recorrido de la solución y un simulador.

.DESCRIPTION
    Toma fuente.html y le inyecta src/1-dominio/dominio.json y las plantillas de correo de
    src/3-adaptadores/power-automate/plantillas-correo, para que el simulador use las mismas reglas
    y los mismos textos que la solución real. Al abrirse, la página comprueba que sus reglas
    reproducen los ejemplos de pruebas/unitarias/Dominio.Tests.ps1.
    Vuelva a ejecutarlo si cambian el dominio o las plantillas.

.EXAMPLE
    ./docs/recorrido/Generar-Recorrido.ps1 -PruebasPester 173 -PruebasPython 14
#>
[CmdletBinding()]
param(
    [int]$PruebasPester = 173,
    [int]$PruebasPython = 14
)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$utf8 = New-Object System.Text.UTF8Encoding($false)

$fuente = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'fuente.html'), $utf8)
$dominio = [System.IO.File]::ReadAllText((Join-Path $raiz 'src\1-dominio\dominio.json'), $utf8).Trim()
$plantillas = [ordered]@{}
foreach ($archivo in Get-ChildItem (Join-Path $raiz 'src\3-adaptadores\power-automate\plantillas-correo') -Filter *.html | Sort-Object Name) {
    $plantillas[$archivo.BaseName] = [System.IO.File]::ReadAllText($archivo.FullName, $utf8).Trim()
}
$meta = [ordered]@{ pester = $PruebasPester; python = $PruebasPython; generado = (Get-Date -Format 'dd/MM/yyyy') }

# "</" se escribe como "<\/" para que ningún texto inyectado cierre el <script> de la página.
$seguro = { param([string]$Json) $Json.Replace('</', '<\/') }
$pagina = $fuente.
    Replace('/*@@DOMINIO@@*/null', (& $seguro $dominio)).
    Replace('/*@@PLANTILLAS@@*/null', (& $seguro ($plantillas | ConvertTo-Json -Compress))).
    Replace('/*@@META@@*/null', (& $seguro ($meta | ConvertTo-Json -Compress)))
if ($pagina -match '/\*@@[A-Z]+@@\*/') { throw "Quedó un marcador sin reemplazar: $($Matches[0])" }

$salida = Join-Path $PSScriptRoot 'recorrido.html'
[System.IO.File]::WriteAllText($salida, $pagina, $utf8)
Write-Host "Recorrido generado: $salida ($([math]::Round($pagina.Length / 1KB)) KB)" -ForegroundColor Green
