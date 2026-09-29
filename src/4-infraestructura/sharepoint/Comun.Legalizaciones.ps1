<#
    Comun.Legalizaciones.ps1 · Capa de INFRAESTRUCTURA (SharePoint)

    Funciones compartidas por 00-Registrar-AppEntraID.ps1, 01-Aprovisionar-SharePoint.ps1
    y 02-Verificar-SharePoint.ps1. No hace cambios por sí mismo.

    Se carga desde los otros scripts con:
        . (Join-Path $PSScriptRoot 'Comun.Legalizaciones.ps1')
#>

# Raíz del proyecto (este archivo está en src/4-infraestructura/sharepoint).
$script:RaizProyecto = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$script:RutaConfigPredeterminada  = Join-Path $script:RaizProyecto 'config\Config.Legalizaciones.psd1'
$script:RutaEsquemaPredeterminada = Join-Path $script:RaizProyecto 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1'
$script:RutaDominioPredeterminada = Join-Path $script:RaizProyecto 'src\1-dominio\dominio.json'

# Tipos de columna del esquema -> tipo de SharePoint (propiedad TypeAsString)
$script:MapaTipos = @{
    Texto           = 'Text'
    TextoMultilinea = 'Note'
    FechaHora       = 'DateTime'
    Fecha           = 'DateTime'
    Opcion          = 'Choice'
    SiNo            = 'Boolean'
    Hipervinculo    = 'URL'
    Persona         = 'User'
}

$script:Contadores = [ordered]@{ 'CREADO' = 0; 'ACTUALIZADO' = 0; 'SIN CAMBIOS' = 0; 'AVISO' = 0 }

# Perfil regional por país (Config.Pais). Ninguno de los dos países tiene horario de verano,
# por eso el desfase UTC de Excel es fijo.
$script:PerfilesRegion = @{
    Mexico   = @{ ZonaSharePoint = 'Guadalajara'; LocaleId = 2058; ZonaWindows = 'Central Standard Time (Mexico)'; DesfaseHorasUTC = -6; Cultura = 'es-MX' }
    Colombia = @{ ZonaSharePoint = 'Quito';       LocaleId = 9226; ZonaWindows = 'SA Pacific Standard Time';       DesfaseHorasUTC = -5; Cultura = 'es-CO' }
}

#region Salida por pantalla ----------------------------------------------------

function Write-Paso {
    param([Parameter(Mandatory)][string]$Texto)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host " $Texto" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Write-Resultado {
    param(
        [Parameter(Mandatory)][ValidateSet('CREADO', 'ACTUALIZADO', 'SIN CAMBIOS', 'AVISO', 'INFO', 'ERROR')][string]$Estado,
        [Parameter(Mandatory)][string]$Texto
    )
    $color = switch ($Estado) {
        'CREADO'      { 'Green' }
        'ACTUALIZADO' { 'Yellow' }
        'SIN CAMBIOS' { 'Gray' }
        'AVISO'       { 'Magenta' }
        'ERROR'       { 'Red' }
        default       { 'White' }
    }
    if ($script:Contadores.Contains($Estado)) { $script:Contadores[$Estado]++ }
    Write-Host ('  [{0,-11}] {1}' -f $Estado, $Texto) -ForegroundColor $color
}

#endregion

#region Configuración -----------------------------------------------------------

# Devuelve la ruta de cada valor de texto que todavía tiene un marcador < >.
function Find-Marcadores {
    param([AllowNull()]$Valor, [string]$Ruta)
    if ($null -eq $Valor) { return }
    if ($Valor -is [System.Collections.IDictionary]) {
        foreach ($clave in $Valor.Keys) { Find-Marcadores -Valor $Valor[$clave] -Ruta ('{0}.{1}' -f $Ruta, $clave) }
    }
    elseif ($Valor -is [string]) {
        if ($Valor -match '<[^<>]*>') { $Ruta }
    }
    elseif ($Valor -is [System.Collections.IEnumerable]) {
        $i = 0
        foreach ($elemento in $Valor) {
            Find-Marcadores -Valor $elemento -Ruta ('{0}[{1}]' -f $Ruta, $i)
            $i++
        }
    }
}

# Valida una persona o grupo de la configuración. Devuelve el error o $null.
function Test-Principal {
    param([AllowNull()]$Principal, [Parameter(Mandatory)][string]$Ruta)
    if ($Principal -isnot [System.Collections.IDictionary]) {
        return "$Ruta debe tener la forma @{ Tipo = '...'; Valor = '...' }."
    }
    $valor = [string]$Principal.Valor
    switch ([string]$Principal.Tipo) {
        'Usuario' {
            if ($valor -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return "$Ruta.Valor debe ser un correo válido (Tipo = 'Usuario')." }
        }
        { $_ -in 'GrupoSeguridad', 'GrupoM365' } {
            $guid = [guid]::Empty
            if (-not [guid]::TryParse($valor, [ref]$guid)) { return "$Ruta.Valor debe ser el Id. de objeto (GUID) del grupo en Entra ID." }
        }
        default { return "$Ruta.Tipo debe ser 'Usuario', 'GrupoSeguridad' o 'GrupoM365'." }
    }
    return $null
}

# Detiene la ejecución si en las secciones indicadas de la configuración quedan marcadores < >.
function Assert-SinMarcadores {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Config,
        [Parameter(Mandatory)][string[]]$Secciones,
        [string]$RutaConfig = 'la configuración'
    )
    $marcadores = @()
    foreach ($seccion in $Secciones) { $marcadores += @(Find-Marcadores -Valor $Config[$seccion] -Ruta ('Config.{0}' -f $seccion)) }
    if ($marcadores.Count -gt 0) {
        throw ("Complete estos valores en '{0}' (todavía tienen un marcador < >):`n  - {1}" -f $RutaConfig, ($marcadores -join "`n  - "))
    }
}

# Perfil regional a partir de Config.Pais (con las excepciones opcionales de Config.Sitio).
function Get-RegionEntorno {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Config)
    $pais = [string]$Config.Pais
    if (-not $script:PerfilesRegion.ContainsKey($pais)) {
        throw ("Config.Pais debe ser uno de: {0}." -f (($script:PerfilesRegion.Keys | Sort-Object) -join ', '))
    }
    $perfil = $script:PerfilesRegion[$pais]
    $zona = $perfil.ZonaSharePoint
    $locale = $perfil.LocaleId
    if ($Config.Sitio -and $Config.Sitio.ZonaHoraria) { $zona = [string]$Config.Sitio.ZonaHoraria }
    if ($Config.Sitio -and $Config.Sitio.LocaleId) { $locale = [int]$Config.Sitio.LocaleId }
    [pscustomobject]@{
        Pais            = $pais
        ZonaSharePoint  = $zona
        LocaleId        = $locale
        ZonaWindows     = $perfil.ZonaWindows
        DesfaseHorasUTC = $perfil.DesfaseHorasUTC
        Cultura         = $perfil.Cultura
    }
}

# Valida las columnas de una sección del esquema (Lista.Columnas o Biblioteca.Columnas).
function Test-ColumnasEsquema {
    param([AllowEmptyCollection()][object[]]$Columnas, [string]$Seccion, $Errores)
    $nombres = @{}
    foreach ($columna in @($Columnas)) {
        $nombre = [string]$columna.NombreInterno
        if ($nombre -notmatch '^[A-Za-z][A-Za-z0-9]{0,31}$') {
            $Errores.Add("Esquema ($Seccion): NombreInterno '$nombre' debe empezar por letra y tener solo letras y números, sin tildes ni espacios (máx. 32).")
        }
        elseif ($nombre -eq 'Title') {
            $Errores.Add("Esquema ($Seccion): 'Title' se configura en Lista.ColumnaTitulo, no como columna.")
        }
        elseif ($nombres.ContainsKey($nombre)) {
            $Errores.Add("Esquema ($Seccion): la columna '$nombre' está repetida.")
        }
        $nombres[$nombre] = $true
        if (-not $script:MapaTipos.ContainsKey([string]$columna.Tipo)) {
            $Errores.Add("Esquema ($Seccion): la columna '$nombre' tiene Tipo '$($columna.Tipo)'. Tipos válidos: $(($script:MapaTipos.Keys | Sort-Object) -join ', ').")
        }
        if ($columna.Tipo -eq 'Opcion') {
            $opciones = @($columna.Opciones)
            if ($opciones.Count -eq 0) { $Errores.Add("Esquema ($Seccion): la columna '$nombre' es de tipo Opcion y no tiene Opciones.") }
            if ($columna.ValorPredeterminado -and $opciones -cnotcontains $columna.ValorPredeterminado) {
                $Errores.Add("Esquema ($Seccion): el ValorPredeterminado de '$nombre' no está entre sus Opciones.")
            }
        }
    }
}

# Carga la configuración del entorno y el esquema, los valida contra el dominio y calcula las URL.
function Import-ConfigLegalizaciones {
    param(
        [Parameter(Mandatory)][string]$RutaConfig,
        [Parameter(Mandatory)][string]$RutaEsquema,
        [string]$RutaDominio = $script:RutaDominioPredeterminada
    )
    foreach ($ruta in @($RutaConfig, $RutaEsquema)) {
        if (-not (Test-Path -LiteralPath $ruta)) { throw "No se encontró el archivo '$ruta'." }
    }
    $cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig
    $esq = Import-PowerShellDataFile -LiteralPath $RutaEsquema

    # 1. Valores sin completar en las secciones que usa SharePoint (el esquema contiene CAML con < >;
    #    la sección Flujos la revisa el generador de configuración de los flujos).
    Assert-SinMarcadores -Config $cfg -Secciones @('Tenant', 'Autenticacion', 'Sitio', 'Permisos', 'Registro') -RutaConfig $RutaConfig

    # 2. Formato de la configuración
    $errores = New-Object System.Collections.Generic.List[string]
    $region = $null
    try { $region = Get-RegionEntorno -Config $cfg } catch { $errores.Add($_.Exception.Message) }
    if ([string]$cfg.Tenant.Nombre -notmatch '^[A-Za-z0-9-]+$') {
        $errores.Add("Tenant.Nombre debe ser solo el prefijo del tenant (ej. 'emobility' para https://emobility.sharepoint.com).")
    }
    $guid = [guid]::Empty
    if (-not [guid]::TryParse([string]$cfg.Autenticacion.ClientId, [ref]$guid)) {
        $errores.Add('Autenticacion.ClientId debe ser el Id. de aplicación (GUID) que devolvió 00-Registrar-AppEntraID.ps1.')
    }
    if ([string]$cfg.Sitio.RutaRelativa -notmatch '^/(sites|teams)/[A-Za-z0-9_-]+$') {
        $errores.Add("Sitio.RutaRelativa debe tener la forma '/sites/NombreSitio' (sin espacios ni tildes).")
    }
    if (@('Sitio', 'ListaYBiblioteca') -notcontains $cfg.Permisos.Ambito) {
        $errores.Add("Permisos.Ambito debe ser 'Sitio' o 'ListaYBiblioteca'.")
    }
    foreach ($grupo in 'PropietariosTI', 'Contabilidad') {
        $principales = @($cfg.Permisos[$grupo])
        if ($principales.Count -eq 0) { $errores.Add("Permisos.$grupo debe tener al menos un elemento.") }
        for ($i = 0; $i -lt $principales.Count; $i++) {
            $problema = Test-Principal -Principal $principales[$i] -Ruta ('Permisos.{0}[{1}]' -f $grupo, $i)
            if ($problema) { $errores.Add($problema) }
        }
    }
    $cuenta = [string]$cfg.Permisos.CuentaServicioFlujo
    if ($cuenta -and $cuenta -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $errores.Add("Permisos.CuentaServicioFlujo debe ser un correo o '' (vacío).")
    }

    # 3. Esquema
    Test-ColumnasEsquema -Columnas @($esq.Lista.Columnas) -Seccion 'Lista' -Errores $errores
    Test-ColumnasEsquema -Columnas @($esq.Biblioteca.Columnas) -Seccion 'Biblioteca' -Errores $errores

    # 4. Regla de dependencia: el esquema de SharePoint no puede contradecir al dominio.
    if ($RutaDominio -and (Test-Path -LiteralPath $RutaDominio)) {
        $dominio = [System.IO.File]::ReadAllText($RutaDominio, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        $estado = @($esq.Lista.Columnas) | Where-Object { $_.NombreInterno -eq 'Estado' } | Select-Object -First 1
        if ($estado) {
            if ((@($estado.Opciones) -join '|') -cne (@($dominio.caso.estados) -join '|')) {
                $errores.Add("Esquema: las opciones de Estado [$(@($estado.Opciones) -join ', ')] no coinciden con los estados del dominio [$(@($dominio.caso.estados) -join ', ')]. Cambie primero src/1-dominio/dominio.json.")
            }
            if ([string]$estado.ValorPredeterminado -cne [string]$dominio.caso.estadoInicial) {
                $errores.Add("Esquema: el valor predeterminado de Estado debe ser el estado inicial del dominio ('$($dominio.caso.estadoInicial)').")
            }
        }
        $raices = @($dominio.carpetas.raizPorEstado.PSObject.Properties | ForEach-Object { [string]$_.Value } | Select-Object -Unique)
        foreach ($raizCarpeta in $raices) {
            if (@($esq.Biblioteca.CarpetasRaiz) -notcontains $raizCarpeta) {
                $errores.Add("Esquema: falta la carpeta raíz '$raizCarpeta' que el dominio usa en carpetas.raizPorEstado.")
            }
        }
    }

    if ($errores.Count -gt 0) {
        throw ("La configuración tiene errores:`n  - " + ($errores -join "`n  - "))
    }

    $tenant = ([string]$cfg.Tenant.Nombre).ToLowerInvariant()
    $raiz = 'https://{0}.sharepoint.com' -f $tenant
    [pscustomobject]@{
        Config  = $cfg
        Esquema = $esq
        Region  = $region
        Urls    = [pscustomobject]@{
            Raiz  = $raiz
            Admin = 'https://{0}-admin.sharepoint.com' -f $tenant
            Sitio = $raiz + $cfg.Sitio.RutaRelativa
        }
    }
}

# Inicia la transcripción de la ejecución y devuelve la ruta del archivo.
function Start-RegistroEjecucion {
    param([Parameter(Mandatory)]$Contexto, [Parameter(Mandatory)][string]$Prefijo)
    $carpeta = [string]$Contexto.Config.Registro.CarpetaLogs
    if (-not $carpeta) { $carpeta = 'logs' }
    if (-not [System.IO.Path]::IsPathRooted($carpeta)) { $carpeta = Join-Path $script:RaizProyecto $carpeta }
    New-Item -ItemType Directory -Path $carpeta -Force | Out-Null
    $ruta = Join-Path $carpeta ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Prefijo, (Get-Date))
    Start-Transcript -Path $ruta | Out-Null
    return $ruta
}

#endregion

#region SharePoint --------------------------------------------------------------

function Connect-SitioLegalizaciones {
    param([Parameter(Mandatory)]$Contexto)
    Write-Resultado 'INFO' ('Conectando a {0} (inicie sesión en la ventana del navegador)' -f $Contexto.Urls.Sitio)
    Connect-PnPOnline -Url $Contexto.Urls.Sitio -Interactive -ClientId $Contexto.Config.Autenticacion.ClientId
}

# Nombre del nivel de permiso por su tipo. Los nombres cambian con el idioma del
# sitio ("Colaborar" / "Contribute"); el tipo (RoleTypeKind) no.
function Get-NombreRol {
    param([Parameter(Mandatory)][ValidateSet('Administrator', 'Editor', 'Contributor', 'Reader')][string]$Tipo)
    $rol = Get-PnPRoleDefinition | Where-Object { [string]$_.RoleTypeKind -eq $Tipo } | Select-Object -First 1
    if (-not $rol) { throw "El sitio no tiene un nivel de permiso de tipo '$Tipo'." }
    return $rol.Name
}

# Tipos de nivel de permiso (RoleTypeKind) que tiene un grupo de SharePoint en el sitio.
function Get-RolesGrupo {
    param([Parameter(Mandatory)]$Grupo)
    $tipos = @()
    try {
        foreach ($rol in (Get-PnPGroupPermissions -Identity $Grupo.Id)) { $tipos += [string]$rol.RoleTypeKind }
    }
    catch {
        # El grupo no tiene permisos asignados en el sitio.
    }
    return $tipos
}

# Nombre de inicio de sesión que SharePoint entiende para agregar a un grupo.
function Get-LoginParaAgregar {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Principal)
    $valor = ([string]$Principal.Valor).Trim()
    switch ([string]$Principal.Tipo) {
        'Usuario'        { return $valor }
        'GrupoSeguridad' { return 'c:0t.c|tenant|{0}' -f $valor.ToLowerInvariant() }
        'GrupoM365'      { return 'c:0o.c|federateddirectoryclaimprovider|{0}' -f $valor.ToLowerInvariant() }
        default          { throw "Tipo de principal no válido: '$($Principal.Tipo)'." }
    }
}

# ¿El principal ya está entre los miembros de un grupo de SharePoint?
function Test-EsMiembro {
    param([AllowNull()][AllowEmptyCollection()][object[]]$Miembros, [Parameter(Mandatory)][System.Collections.IDictionary]$Principal)
    $valor = ([string]$Principal.Valor).Trim()
    if ($Principal.Tipo -eq 'Usuario') {
        return [bool](@($Miembros) | Where-Object { $_ -and ($_.Email -eq $valor -or $_.LoginName -like "*|$valor") })
    }
    $claim = Get-LoginParaAgregar -Principal $Principal
    return [bool](@($Miembros) | Where-Object { $_ -and $_.LoginName -eq $claim })
}

function ConvertTo-XmlSeguro {
    param([AllowNull()][AllowEmptyString()][string]$Texto)
    return [System.Security.SecurityElement]::Escape([string]$Texto)
}

# Definición CAML de una columna del esquema.
# Se crea con DisplayName = NombreInterno para que SharePoint use exactamente ese
# nombre interno (sin tildes ni espacios); luego el script le pone el nombre visible.
function New-FieldXml {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Columna)
    $nombre = [string]$Columna.NombreInterno
    $requerida = if ($Columna.Requerida) { 'TRUE' } else { 'FALSE' }
    $descripcion = ConvertTo-XmlSeguro $Columna.Descripcion
    $base = 'Name="{0}" StaticName="{0}" DisplayName="{0}" Required="{1}" Description="{2}"' -f $nombre, $requerida, $descripcion
    switch ([string]$Columna.Tipo) {
        'Texto'           { return '<Field Type="Text" {0} MaxLength="255" />' -f $base }
        'TextoMultilinea' { return '<Field Type="Note" {0} NumLines="6" RichText="FALSE" AppendOnly="FALSE" />' -f $base }
        'FechaHora'       { return '<Field Type="DateTime" {0} Format="DateTime" FriendlyDisplayFormat="Disabled" />' -f $base }
        'Fecha'           { return '<Field Type="DateTime" {0} Format="DateOnly" FriendlyDisplayFormat="Disabled" />' -f $base }
        'SiNo'            { return '<Field Type="Boolean" {0}><Default>0</Default></Field>' -f $base }
        'Hipervinculo'    { return '<Field Type="URL" {0} Format="Hyperlink" />' -f $base }
        'Persona'         { return '<Field Type="User" {0} UserSelectionMode="PeopleOnly" UserSelectionScope="0" />' -f $base }
        'Opcion' {
            $opciones = (@($Columna.Opciones) | ForEach-Object { '<CHOICE>{0}</CHOICE>' -f (ConvertTo-XmlSeguro $_) }) -join ''
            $predeterminado = ''
            if ($Columna.ValorPredeterminado) { $predeterminado = '<Default>{0}</Default>' -f (ConvertTo-XmlSeguro $Columna.ValorPredeterminado) }
            return '<Field Type="Choice" {0} Format="Dropdown" FillInChoice="FALSE">{1}<CHOICES>{2}</CHOICES></Field>' -f $base, $predeterminado, $opciones
        }
        default { throw "Tipo de columna no soportado: '$($Columna.Tipo)'." }
    }
}

# Lee de la definición XML de una columna su formato, opciones y valor predeterminado.
function Get-DefinicionCampo {
    param([Parameter(Mandatory)]$Campo)
    $xml = [xml]$Campo.SchemaXml
    $opciones = @()
    if ($xml.Field.CHOICES) { $opciones = @($xml.Field.CHOICES.CHOICE | ForEach-Object { [string]$_ }) }
    [pscustomobject]@{
        Formato        = [string]$xml.Field.Format
        Opciones       = $opciones
        Predeterminado = [string]$xml.Field.Default
    }
}

# Normaliza CAML para comparar consultas sin que importen espacios o comillas.
function Format-Caml {
    param([AllowNull()][AllowEmptyString()][string]$Caml)
    return (($Caml -replace "'", '"') -replace '\s+', '')
}

#endregion
