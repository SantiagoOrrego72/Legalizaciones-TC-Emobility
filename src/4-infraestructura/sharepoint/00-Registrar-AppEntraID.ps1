#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' }
<#
.SYNOPSIS
    Registra en Entra ID la aplicación con la que PnP PowerShell inicia sesión.
    Se ejecuta UNA sola vez por tenant.

.DESCRIPTION
    Desde septiembre de 2024 PnP PowerShell exige que cada organización use su propia
    aplicación registrada en Entra ID. Este script la crea con permisos DELEGADOS:

      - SharePoint:      AllSites.FullControl
      - Microsoft Graph: User.Read

    "Delegados" significa que la aplicación actúa con los permisos de la persona que
    inicia sesión: no le da a nadie acceso a sitios que no tenga ya.

    Quién lo ejecuta: alguien con rol Administrador de aplicaciones en la nube o
    Administrador global en Entra ID. Al final se abre el navegador para aceptar el
    consentimiento en nombre de la organización.

    Si la organización ya tiene una aplicación de PnP PowerShell, NO ejecute este script:
    pida su Id. de aplicación y póngalo en Config.Legalizaciones.psd1 -> Autenticacion.ClientId.

.PARAMETER RutaConfig
    Archivo de configuración. Solo se usan Tenant.DominioEntra y Autenticacion.NombreApp.

.EXAMPLE
    ./00-Registrar-AppEntraID.ps1
#>
[CmdletBinding()]
param(
    [string]$RutaConfig = (Join-Path $PSScriptRoot '..\..\..\config\Config.Legalizaciones.psd1')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Comun.Legalizaciones.ps1')

if (-not (Test-Path -LiteralPath $RutaConfig)) { throw "No se encontró el archivo '$RutaConfig'." }
$cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig

# En este punto el ClientId todavía no existe: solo se validan los dos valores que se usan.
$faltantes = @(Find-Marcadores -Valor $cfg.Tenant.DominioEntra -Ruta 'Config.Tenant.DominioEntra') +
             @(Find-Marcadores -Valor $cfg.Autenticacion.NombreApp -Ruta 'Config.Autenticacion.NombreApp')
if ($faltantes.Count -gt 0) {
    throw ("Complete estos valores en '{0}':`n  - {1}" -f $RutaConfig, ($faltantes -join "`n  - "))
}
if ([string]$cfg.Tenant.DominioEntra -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') {
    throw "Tenant.DominioEntra debe ser un dominio, por ejemplo 'emobility.onmicrosoft.com'."
}

Write-Paso 'Registro de la aplicación de Entra ID para PnP PowerShell'
Write-Resultado 'INFO' "Tenant: $($cfg.Tenant.DominioEntra)"
Write-Resultado 'INFO' "Aplicación: $($cfg.Autenticacion.NombreApp)"
Write-Resultado 'INFO' 'Se abrirá el navegador: inicie sesión y acepte el consentimiento en nombre de la organización.'

$resultado = Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName $cfg.Autenticacion.NombreApp `
    -Tenant $cfg.Tenant.DominioEntra `
    -SharePointDelegatePermissions 'AllSites.FullControl' `
    -GraphDelegatePermissions 'User.Read'

Write-Paso 'Siguiente paso'
Write-Host (($resultado | Format-List | Out-String).Trim())
Write-Host ''
Write-Host '  1. Copie el Id. de aplicación (AzureAppId / ClientId) que aparece arriba en' -ForegroundColor Yellow
Write-Host '     config\Config.Legalizaciones.psd1 -> Autenticacion.ClientId' -ForegroundColor Yellow
Write-Host '  2. Complete el resto de config\Config.Legalizaciones.psd1' -ForegroundColor Yellow
Write-Host '  3. Ejecute ./src/4-infraestructura/sharepoint/01-Aprovisionar-SharePoint.ps1' -ForegroundColor Yellow
