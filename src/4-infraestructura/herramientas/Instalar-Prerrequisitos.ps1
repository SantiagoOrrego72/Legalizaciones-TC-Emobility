<#
.SYNOPSIS
    Instala en este equipo las herramientas que usa el proyecto. No requiere permisos de administrador.

.DESCRIPTION
    1. PowerShell 7.4 o superior (lo exige PnP.PowerShell):
         - Si ya existe, no hace nada.
         - Con permisos de administrador y winget: winget install Microsoft.PowerShell.
         - Sin permisos de administrador: versión portátil oficial (ZIP publicado por Microsoft en
           github.com/PowerShell/PowerShell), verificada con SHA256 y extraída en
           %LOCALAPPDATA%\Programs\PowerShell-7.
    2. Módulos de PowerShell 7 para el usuario actual (PowerShell Gallery):
         - PnP.PowerShell 3.x (aprovisionamiento y validación de SharePoint).
         - Pester 5.x (pruebas automáticas).
    3. Paquete openpyxl de Python (PyPI), si Python está instalado: plantilla del comprobante y datos de prueba.

    Perfiles:
      - Administracion: solo lo necesario para instalar y administrar la solución en el equipo
        de la empresa (PowerShell 7 y PnP.PowerShell). Recomendado para el equipo de producción.
      - Completo (predeterminado): además Pester y openpyxl, para ejecutar las pruebas automáticas
        y regenerar la plantilla o los datos de prueba (equipo de desarrollo o mantenimiento).

    Se ejecuta con Windows PowerShell 5.1 (el que trae Windows) o con PowerShell 7.
    Se puede ejecutar varias veces: lo que ya está instalado no se vuelve a instalar.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\src\4-infraestructura\herramientas\Instalar-Prerrequisitos.ps1 -Perfil Administracion
#>
[CmdletBinding()]
param(
    [ValidateSet('Completo', 'Administracion')][string]$Perfil = 'Completo',
    # No instala openpyxl.
    [switch]$OmitirPython
)
if ($Perfil -eq 'Administracion') { $OmitirPython = $true }

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # sin barra de progreso las descargas son mucho más rápidas en 5.1
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$carpetaPortable = Join-Path $env:LOCALAPPDATA 'Programs\PowerShell-7'
$versionMinima = [version]'7.4.0'

function Get-VersionPwsh {
    param([string]$Ruta)
    try {
        $texto = & $Ruta -NoProfile -NoLogo -Command '$PSVersionTable.PSVersion.ToString()'
        return [version](($texto | Select-Object -First 1) -replace '-.*$', '')
    }
    catch { return $null }
}

# Devuelve la ruta de un pwsh.exe 7.4+ (instalado o portátil) o $null.
function Find-Pwsh {
    $candidatos = New-Object System.Collections.Generic.List[string]
    $comando = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($comando) { $candidatos.Add($comando.Source) }
    $candidatos.Add((Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))
    $candidatos.Add((Join-Path $carpetaPortable 'pwsh.exe'))
    foreach ($ruta in $candidatos) {
        if ((Test-Path $ruta) -and ((Get-VersionPwsh $ruta) -ge $versionMinima)) { return $ruta }
    }
    return $null
}

# ---------------------------------------------------------------- 1. PowerShell 7
Write-Host '1/3 PowerShell 7' -ForegroundColor Cyan
$pwsh = Find-Pwsh
if ($pwsh) {
    Write-Host "  Ya disponible: $pwsh ($(Get-VersionPwsh $pwsh))"
}
else {
    $esAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($esAdmin -and (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host '  Instalando con winget (Microsoft.PowerShell)...'
        winget install --id Microsoft.PowerShell --source winget --silent
    }
    else {
        Write-Host '  Sin permisos de administrador: se instala la versión portátil oficial.'
        $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -UseBasicParsing
        $version = $release.tag_name.TrimStart('v')
        $nombreZip = "PowerShell-$version-win-x64.zip"
        $asset = $release.assets | Where-Object { $_.name -eq $nombreZip } | Select-Object -First 1
        if (-not $asset) { throw "No se encontró $nombreZip en la versión $($release.tag_name)." }

        $zip = Join-Path $env:TEMP $nombreZip
        Write-Host ("  Descargando {0} ({1:N1} MB) desde {2}" -f $nombreZip, ($asset.size / 1MB), $asset.browser_download_url)
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -UseBasicParsing

        # Verificación de integridad con el SHA256 publicado por Microsoft.
        $esperado = $null
        if ($asset.PSObject.Properties['digest'] -and $asset.digest -like 'sha256:*') { $esperado = $asset.digest.Substring(7) }
        if (-not $esperado) {
            $hashes = $release.assets | Where-Object { $_.name -eq 'hashes.sha256' } | Select-Object -First 1
            if ($hashes) {
                $archivoHashes = Join-Path $env:TEMP 'powershell-hashes.sha256'
                Invoke-WebRequest -Uri $hashes.browser_download_url -OutFile $archivoHashes -UseBasicParsing
                $linea = Get-Content $archivoHashes | Where-Object { $_ -match [regex]::Escape($nombreZip) } | Select-Object -First 1
                if ($linea) { $esperado = ($linea -split '\s+')[0] }
            }
        }
        $calculado = (Get-FileHash -Path $zip -Algorithm SHA256).Hash
        if ($esperado) {
            if ($calculado -ne $esperado.ToUpperInvariant()) { throw "El SHA256 de $nombreZip no coincide con el publicado. No se instala." }
            Write-Host "  SHA256 verificado: $calculado"
        }
        else {
            Write-Warning "No se encontró el SHA256 publicado; SHA256 calculado: $calculado"
        }

        if (Test-Path $carpetaPortable) { Remove-Item $carpetaPortable -Recurse -Force }
        Expand-Archive -Path $zip -DestinationPath $carpetaPortable -Force
        Remove-Item $zip -Force
        Write-Host "  Instalado en $carpetaPortable"
    }
    $pwsh = Find-Pwsh
    if (-not $pwsh) { throw 'No quedó disponible PowerShell 7.4 o superior.' }
    Write-Host "  Listo: $pwsh ($(Get-VersionPwsh $pwsh))"
}

# ---------------------------------------------------------------- 2. Módulos
Write-Host '2/3 Módulos de PowerShell 7 (PowerShell Gallery)' -ForegroundColor Cyan
$scriptModulos = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$modulos = @(
    @{ Nombre = 'PnP.PowerShell'; Minima = [version]'3.0.0' }
    @{ Nombre = 'Pester';         Minima = [version]'5.5.0' }
)
if ('__PERFIL__' -eq 'Administracion') { $modulos = @($modulos[0]) }
foreach ($m in $modulos) {
    $instalado = Get-Module -ListAvailable -Name $m.Nombre | Where-Object { $_.Version -ge $m.Minima } | Sort-Object Version -Descending | Select-Object -First 1
    if ($instalado) { Write-Host ("  {0} {1} ya instalado" -f $m.Nombre, $instalado.Version); continue }
    Write-Host ("  Instalando {0}..." -f $m.Nombre)
    if (Get-Command Install-PSResource -ErrorAction SilentlyContinue) {
        Install-PSResource -Name $m.Nombre -Scope CurrentUser -TrustRepository -Quiet
    }
    else {
        Install-Module -Name $m.Nombre -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
    $instalado = Get-Module -ListAvailable -Name $m.Nombre | Sort-Object Version -Descending | Select-Object -First 1
    Write-Host ("  {0} {1} instalado" -f $m.Nombre, $instalado.Version)
}
'@
# Se envía codificado: Windows PowerShell 5.1 pierde las comillas al pasar texto a otro ejecutable.
$comandoCodificado = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptModulos.Replace('__PERFIL__', $Perfil)))
& $pwsh -NoProfile -NoLogo -EncodedCommand $comandoCodificado
if ($LASTEXITCODE -ne 0) { throw 'Falló la instalación de módulos de PowerShell 7.' }

# ---------------------------------------------------------------- 3. Python
Write-Host '3/3 Python: openpyxl' -ForegroundColor Cyan
if ($OmitirPython) {
    Write-Host "  Omitido (perfil $Perfil)."
}
else {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
    if (-not $python -or $python.Source -like '*WindowsApps*') {
        Write-Warning 'No se encontró Python. Instálelo desde https://www.python.org/downloads/ y vuelva a ejecutar este script (solo se necesita para la plantilla del comprobante y los datos de prueba).'
    }
    else {
        & $python.Source -m pip install --upgrade --disable-pip-version-check --quiet openpyxl
        if ($LASTEXITCODE -ne 0) { throw 'Falló pip install openpyxl.' }
        $versionOpenpyxl = & $python.Source -c "import openpyxl; print(openpyxl.__version__)"
        Write-Host "  openpyxl $versionOpenpyxl ($($python.Source))"
    }
}

Write-Host ''
Write-Host 'Prerrequisitos listos.' -ForegroundColor Green
Write-Host "PowerShell 7: $pwsh"
Write-Host 'Para trabajar en el proyecto abra Abrir-PowerShell7.cmd (en la raíz del proyecto).'
