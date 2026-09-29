<#
.SYNOPSIS
    Genera el JSON que se pega en la acción "Configuracion" de los flujos LEG-01 y LEG-02.

.DESCRIPTION
    Capa de INFRAESTRUCTURA (Power Automate). Une en un solo JSON:
      - las reglas de negocio de src/1-dominio/dominio.json (estados, extensiones, ventana de
        duplicados, mensajes, formato del ID),
      - el contrato de SharePoint (biblioteca, carpetas, lista) de Esquema.Legalizaciones.psd1,
      - los datos del entorno de config/Config.Legalizaciones.psd1 (URL del sitio, buzón,
        correos de aviso, zona horaria del país).
    Así los flujos no tienen reglas escritas a mano: si cambia el dominio, se vuelve a generar
    este JSON y se reemplaza el contenido de la acción "Configuracion" en ambos flujos.

    Funciona en Windows PowerShell 5.1 y PowerShell 7. No se conecta a nada.

.PARAMETER Borrador
    Genera el archivo aunque la configuración tenga marcadores < > sin completar
    (salida/Configuracion-Flujos.BORRADOR.json). Sirve para revisarlo, NO para pegarlo en los flujos.

.EXAMPLE
    ./src/4-infraestructura/power-automate/Generar-ConfiguracionFlujos.ps1
#>
[CmdletBinding()]
param(
    [string]$RutaConfig,
    [string]$RutaSalida,
    [switch]$Borrador
)

$ErrorActionPreference = 'Stop'
$aqui = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$raiz = (Resolve-Path (Join-Path $aqui '..\..\..')).Path
if (-not $RutaConfig) { $RutaConfig = Join-Path $raiz 'config\Config.Legalizaciones.psd1' }
. (Join-Path $raiz 'src\4-infraestructura\sharepoint\Comun.Legalizaciones.ps1')

$cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig
$esq = Import-PowerShellDataFile -LiteralPath (Join-Path $raiz 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
$dominio = [System.IO.File]::ReadAllText((Join-Path $raiz 'src\1-dominio\dominio.json'), [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$region = Get-RegionEntorno -Config $cfg

if (-not $Borrador) {
    Assert-SinMarcadores -Config $cfg -Secciones @('Tenant', 'Sitio', 'Flujos') -RutaConfig $RutaConfig
}
if (-not $RutaSalida) {
    $nombre = if ($Borrador) { 'Configuracion-Flujos.BORRADOR.json' } else { 'Configuracion-Flujos.json' }
    $RutaSalida = Join-Path $raiz "salida\$nombre"
}

$urlSitio = 'https://{0}.sharepoint.com{1}' -f ([string]$cfg.Tenant.Nombre).ToLowerInvariant(), $cfg.Sitio.RutaRelativa
$raices = $dominio.carpetas.raizPorEstado
$extensiones = [ordered]@{}
$permitidas = @()
$comparan = @()
foreach ($tipo in $dominio.documentos.PSObject.Properties) {
    $extensiones[$tipo.Name] = @($tipo.Value.extensiones)
    $permitidas += @($tipo.Value.extensiones)
    if ($tipo.Value.comparaDuplicados) { $comparan += @($tipo.Value.extensiones) }
}
$mensajes = [ordered]@{}
foreach ($mensaje in $dominio.mensajes.PSObject.Properties) { $mensajes[$mensaje.Name] = $mensaje.Value }

$configuracion = [ordered]@{
    version                       = $dominio.version
    prefijoId                     = $dominio.caso.prefijoId
    formatoConsecutivo            = ('0' * [int]$dominio.caso.digitosConsecutivo)
    estadoInicial                 = $dominio.caso.estadoInicial
    estadoAprobado                = $dominio.caso.estadoAprobado
    estadoRechazado               = $dominio.caso.estadoRechazado
    zonaHorariaWindows            = $region.ZonaWindows
    urlSitio                      = $urlSitio
    urlElemento                   = '{0}/{1}/DispForm.aspx?ID=' -f $urlSitio, $esq.Lista.Url
    rutaBiblioteca                = '/' + $esq.Biblioteca.Url
    carpetaInicial                = $raices.($dominio.caso.estadoInicial)
    carpetaAprobadas              = $raices.($dominio.caso.estadoAprobado)
    carpetaRechazadas             = $raices.($dominio.caso.estadoRechazado)
    extensiones                   = $extensiones
    extensionesPermitidas         = $permitidas
    extensionesComparanDuplicados = $comparan
    diasVentanaDuplicado          = [int]$dominio.duplicados.diasVentana
    estadosExcluidosDuplicado     = @($dominio.duplicados.estadosExcluidos)
    largoMaximoCarpeta            = [int]$dominio.carpetas.largoMaximoNombre
    nombreSinDato                 = $dominio.carpetas.nombreSinDato
    prefijoNotaInterna            = $dominio.observaciones.prefijoNotaInterna
    buzonLegalizaciones           = [string]$cfg.Flujos.BuzonLegalizaciones
    correoContabilidad            = [string]$cfg.Flujos.CorreoContabilidad
    correoSoporteTI               = [string]$cfg.Flujos.CorreoSoporteTI
    notificarCierre               = [bool]$cfg.Flujos.NotificarCierre
    mensajes                      = $mensajes
}

New-Item -ItemType Directory -Path (Split-Path $RutaSalida) -Force | Out-Null
$json = $configuracion | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($RutaSalida, $json, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Configuración de los flujos generada: $RutaSalida" -ForegroundColor Green
if ($Borrador) {
    Write-Host 'BORRADOR: puede contener marcadores < >. No lo pegue en los flujos hasta completar la configuración y generarlo sin -Borrador.' -ForegroundColor Yellow
}
else {
    Write-Host 'Pegue su contenido completo en la acción "Configuracion" de LEG-01 y de LEG-02.'
}
