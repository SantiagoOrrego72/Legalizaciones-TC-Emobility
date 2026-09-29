<#
.SYNOPSIS
    Revisa, sin cambiar nada, si este equipo está listo para instalar y administrar la solución.

.DESCRIPTION
    La solución corre en Microsoft 365 (SharePoint, Power Automate, Exchange). Este equipo solo se usa
    para instalarla y mantenerla. Este diagnóstico comprueba:

      Obligatorio (equipo de administración de TI)
        - Windows 10/11
        - PowerShell 7.4 o superior
        - Módulo PnP.PowerShell 3.x (para PowerShell 7)
        - Microsoft Excel de escritorio (Microsoft 365 Apps): genera el libro de análisis
        - Política de ejecución que permita los scripts del proyecto
        - Scripts del proyecto sin bloqueo de "descargado de internet"
        - Acceso HTTPS a los servicios de Microsoft 365 (y a la PowerShell Gallery para instalar módulos)
      Opcional (mantenimiento y pruebas)
        - Pester 5+ (pruebas automáticas), Python 3 + openpyxl (plantilla y datos de prueba),
          Outlook clásico (script de correos de prueba), Git (control de versiones)

    Termina con código 1 si falta algo obligatorio. Funciona en Windows PowerShell 5.1 y PowerShell 7.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\src\4-infraestructura\herramientas\Diagnosticar-Equipo.ps1
#>
[CmdletBinding()]
param(
    [string]$RutaConfig
)
# $PSScriptRoot llega vacío a los valores predeterminados de param() con powershell.exe -File (5.1).
$aqui = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $RutaConfig) { $RutaConfig = Join-Path $aqui '..\..\..\config\Config.Legalizaciones.psd1' }

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$raiz = (Resolve-Path (Join-Path $aqui '..\..\..')).Path
$resultados = New-Object System.Collections.Generic.List[object]

function Add-Chequeo {
    param([string]$Nivel, [string]$Elemento, [bool]$Cumple, [string]$Detalle, [string]$Solucion)
    $estado = if ($Cumple) { 'OK' } elseif ($Nivel -eq 'Obligatorio') { 'FALTA' } else { 'NO INSTALADO' }
    $resultados.Add([pscustomobject]@{
        Nivel = $Nivel; Elemento = $Elemento; Estado = $estado; Detalle = $Detalle
        Solucion = $(if ($Cumple) { '' } else { $Solucion })
    })
}

function Find-Pwsh {
    $candidatos = @()
    $comando = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($comando) { $candidatos += $comando.Source }
    $candidatos += (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe')
    $candidatos += (Join-Path $env:LOCALAPPDATA 'Programs\PowerShell-7\pwsh.exe')
    foreach ($ruta in $candidatos) { if (Test-Path $ruta) { return $ruta } }
    return $null
}

function Invoke-Pwsh([string]$Pwsh, [string]$Comando) {
    $codificado = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Comando))
    return (& $Pwsh -NoProfile -NoLogo -EncodedCommand $codificado 2>$null | Select-Object -First 1)
}

function Test-Url([string]$Url) {
    try {
        Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -TimeoutSec 10 -MaximumRedirection 0 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        # Cualquier respuesta HTTP (redirección, 401, 403, 405) prueba que el servicio es alcanzable.
        $respuesta = $null
        if ($_.Exception.PSObject.Properties['Response']) { $respuesta = $_.Exception.Response }
        return ($null -ne $respuesta)
    }
}

Write-Host 'Diagnóstico del equipo para Legalizaciones TC' -ForegroundColor Cyan
Write-Host "Equipo: $env:COMPUTERNAME · Usuario: $env:USERNAME"
Write-Host ''

# ------------------------------------------------------------------ Sistema
$version = [Environment]::OSVersion.Version
Add-Chequeo 'Obligatorio' 'Windows 10 u 11' ($version.Major -ge 10) ("Windows {0}" -f $version) 'El generador del libro de Excel necesita Windows con Excel de escritorio.'

# ------------------------------------------------------------------ PowerShell 7 y módulos
$pwsh = Find-Pwsh
$versionPwsh = $null
if ($pwsh) { $versionPwsh = Invoke-Pwsh $pwsh '$PSVersionTable.PSVersion.ToString()' }
$pwshOk = $false
if ($versionPwsh) { $pwshOk = [version]($versionPwsh -replace '-.*$', '') -ge [version]'7.4.0' }
Add-Chequeo 'Obligatorio' 'PowerShell 7.4 o superior' $pwshOk $(if ($pwsh) { "$versionPwsh en $pwsh" } else { 'No encontrado' }) 'TI: winget install --id Microsoft.PowerShell (o Intune). Sin permisos: Instalar-Prerrequisitos.cmd instala la versión portátil.'

$pnp = $null; $pester = $null
if ($pwshOk) {
    $pnp = Invoke-Pwsh $pwsh "(Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()"
    $pester = Invoke-Pwsh $pwsh "(Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()"
}
Add-Chequeo 'Obligatorio' 'Módulo PnP.PowerShell 3.x' ([bool]$pnp -and [version]$pnp -ge [version]'3.0.0') $(if ($pnp) { "Versión $pnp" } else { 'No instalado' }) 'En PowerShell 7: Install-Module PnP.PowerShell -Scope CurrentUser (o Instalar-Prerrequisitos.cmd).'
Add-Chequeo 'Opcional' 'Pester 5+ (pruebas automáticas)' ([bool]$pester -and [version]$pester -ge [version]'5.5.0') $(if ($pester) { "Versión $pester" } else { 'No instalado' }) 'Instalar-Prerrequisitos.cmd (perfil Completo).'

# ------------------------------------------------------------------ Office
$excelClsid = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CLSID' -ErrorAction SilentlyContinue).'(default)'
$office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$versionOffice = if ($office) { "$($office.ProductReleaseIds) $($office.VersionToReport) $($office.Platform)" } else { 'versión no detectada' }
Add-Chequeo 'Obligatorio' 'Excel de escritorio (Microsoft 365 Apps)' ([bool]$excelClsid) $(if ($excelClsid) { $versionOffice } else { 'No instalado' }) 'Instalar Microsoft 365 Apps. Se usa para generar el libro de análisis y para actualizar Power Query.'
$outlookClsid = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction SilentlyContinue).'(default)'
Add-Chequeo 'Opcional' 'Outlook clásico (correos de prueba automáticos)' ([bool]$outlookClsid) $(if ($outlookClsid) { 'Registrado' } else { 'No instalado' }) 'Sin Outlook clásico, los correos de prueba se envían a mano.'

# ------------------------------------------------------------------ Python y Git
$python = Get-Command python -ErrorAction SilentlyContinue
if ($python -and $python.Source -like '*WindowsApps*') { $python = $null }
$openpyxl = $null
if ($python) { try { $openpyxl = & $python.Source -c "import openpyxl; print(openpyxl.__version__)" 2>$null } catch { $openpyxl = $null } }
Add-Chequeo 'Opcional' 'Python 3 + openpyxl (plantilla y datos de prueba)' ([bool]$openpyxl) $(if ($openpyxl) { "openpyxl $openpyxl" } elseif ($python) { 'Python sin openpyxl' } else { 'No instalado' }) 'Solo para regenerar la plantilla o los datos de prueba. python -m pip install openpyxl'
$git = Get-Command git -ErrorAction SilentlyContinue
Add-Chequeo 'Opcional' 'Git (control de versiones del proyecto)' ([bool]$git) $(if ($git) { (& git --version) } else { 'No instalado' }) 'Recomendado para guardar el proyecto en el repositorio de TI.'

# ------------------------------------------------------------------ Política de ejecución y bloqueo de archivos
if ($pwshOk) {
    $politica = Invoke-Pwsh $pwsh "(Get-ExecutionPolicy).ToString()"
    $porGpo = Invoke-Pwsh $pwsh "((Get-ExecutionPolicy -List | Where-Object { `$_.Scope -in 'MachinePolicy','UserPolicy' -and `$_.ExecutionPolicy -ne 'Undefined' }).ExecutionPolicy -join ',')"
    $permitida = $politica -in @('RemoteSigned', 'Unrestricted', 'Bypass')
    $detalle = "Efectiva en PowerShell 7: $politica" + $(if ($porGpo) { " (impuesta por directiva: $porGpo)" } else { '' })
    Add-Chequeo 'Obligatorio' 'Política de ejecución de scripts' $permitida $detalle 'Si la impone una directiva (AllSigned), TI debe firmar los scripts con su certificado de firma de código o autorizar RemoteSigned para este equipo.'
}
$bloqueados = @(Get-ChildItem $raiz -Recurse -File -Include *.ps1, *.psm1, *.psd1, *.cmd -ErrorAction SilentlyContinue |
    Where-Object { Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue })
Add-Chequeo 'Obligatorio' 'Scripts sin bloqueo de "descargado de internet"' ($bloqueados.Count -eq 0) $(if ($bloqueados.Count) { "$($bloqueados.Count) archivos bloqueados" } else { 'Ninguno bloqueado' }) 'En PowerShell, desde la carpeta del proyecto: Get-ChildItem -Recurse | Unblock-File'

# ------------------------------------------------------------------ Red
$tenant = $null
if (Test-Path $RutaConfig) {
    try {
        $cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig
        if ([string]$cfg.Tenant.Nombre -match '^[A-Za-z0-9-]+$') { $tenant = ([string]$cfg.Tenant.Nombre).ToLowerInvariant() }
    }
    catch { $tenant = $null }
}
$servicios = [ordered]@{
    'Inicio de sesión de Microsoft (login.microsoftonline.com)' = 'https://login.microsoftonline.com'
    'Microsoft Graph (graph.microsoft.com)'                      = 'https://graph.microsoft.com'
    'Power Automate (make.powerautomate.com)'                    = 'https://make.powerautomate.com'
    'PowerShell Gallery (instalación de módulos)'                = 'https://www.powershellgallery.com'
}
if ($tenant) { $servicios["SharePoint del tenant ($tenant.sharepoint.com)"] = "https://$tenant.sharepoint.com" }
foreach ($nombre in $servicios.Keys) {
    Add-Chequeo 'Obligatorio' "Acceso a $nombre" (Test-Url $servicios[$nombre]) $servicios[$nombre] 'Pida a TI/redes permitir HTTPS (443) a ese destino o configurar el proxy.'
}
if (-not $tenant) {
    $resultados.Add([pscustomobject]@{ Nivel = 'Información'; Elemento = 'Acceso al SharePoint del tenant'; Estado = 'SIN REVISAR'; Detalle = 'Tenant.Nombre no está configurado'; Solucion = 'Complete config/Config.Legalizaciones.psd1 y vuelva a ejecutar el diagnóstico.' })
}

# ------------------------------------------------------------------ Resultado
$resultados | Format-Table Nivel, Elemento, Estado, Detalle -AutoSize -Wrap | Out-Host
$pendientes = @($resultados | Where-Object { $_.Estado -ne 'OK' -and $_.Solucion })
if ($pendientes.Count) {
    Write-Host 'Qué hacer:' -ForegroundColor Yellow
    foreach ($p in $pendientes) { Write-Host ("  - {0} [{1}]: {2}" -f $p.Elemento, $p.Estado, $p.Solucion) }
    Write-Host ''
}
$faltas = @($resultados | Where-Object { $_.Estado -eq 'FALTA' }).Count
if ($faltas -gt 0) {
    Write-Host "RESULTADO: faltan $faltas requisitos obligatorios." -ForegroundColor Red
    exit 1
}
Write-Host 'RESULTADO: el equipo tiene todo lo obligatorio para instalar y administrar la solución.' -ForegroundColor Green
