#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' }
<#
.SYNOPSIS
    Aprovisiona en SharePoint Online la estructura de Legalizaciones TC (Emobility, Fase 1).
    Capa de INFRAESTRUCTURA: aplica el contrato de datos del adaptador de SharePoint.

.DESCRIPTION
    Crea o corrige, en este orden:
      1. Sitio dedicado sin Grupo de Microsoft 365 (plantilla STS#3), si Sitio.Crear = $true.
      2. Configuración regional del sitio (zona horaria y formato de fecha y número).
      3. Biblioteca Legalizaciones_TC: carpetas raíz Pendientes / Aprobadas / Rechazadas,
         columna ID Legalización y vista predeterminada.
      4. Lista Registro_Legalizaciones_TC (registro operativo).
      5. Columnas del registro, según Esquema.Legalizaciones.psd1.
      6. Vistas de trabajo de Contabilidad.
      7. Permisos: TI con Control total; Contabilidad y la cuenta del flujo con Colaborar.
      8. Restricciones de uso compartido.

    Se puede ejecutar las veces que haga falta: lo que ya está como indica la
    configuración no se toca y lo que difiere se corrige. Nunca borra elementos,
    archivos, columnas ni listas, y no cambia el tipo de una columna existente.

.PARAMETER RutaConfig
    Configuración del entorno. Por defecto, config/Config.Legalizaciones.psd1 del proyecto.

.PARAMETER RutaEsquema
    Contrato de datos (biblioteca, lista, columnas, vistas). Por defecto,
    src/3-adaptadores/sharepoint/Esquema.Legalizaciones.psd1.

.EXAMPLE
    ./src/4-infraestructura/sharepoint/01-Aprovisionar-SharePoint.ps1

.EXAMPLE
    ./src/4-infraestructura/sharepoint/01-Aprovisionar-SharePoint.ps1 -RutaConfig ./config/Config.Pruebas.psd1
#>
[CmdletBinding()]
param(
    [string]$RutaConfig  = (Join-Path $PSScriptRoot '..\..\..\config\Config.Legalizaciones.psd1'),
    [string]$RutaEsquema = (Join-Path $PSScriptRoot '..\..\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Comun.Legalizaciones.ps1')

#region Funciones de este script -------------------------------------------------

# Aplica a una lista o biblioteca solo las propiedades que difieren de lo deseado.
# Las claves de $Deseado son parámetros de Set-PnPList.
function Sync-PropiedadesLista {
    param(
        [Parameter(Mandatory)][string]$Titulo,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Deseado
    )
    $propiedadEquivalente = @{ MajorVersions = 'MajorVersionLimit' }
    $lista = Get-PnPList -Identity $Titulo -Includes Description, EnableVersioning, EnableMinorVersions, MajorVersionLimit, ForceCheckout, EnableAttachments, EnableFolderCreation
    $cambios = @{}
    foreach ($parametro in $Deseado.Keys) {
        $propiedad = $parametro
        if ($propiedadEquivalente.ContainsKey($parametro)) { $propiedad = $propiedadEquivalente[$parametro] }
        if ([string]$lista.$propiedad -cne [string]$Deseado[$parametro]) { $cambios[$parametro] = $Deseado[$parametro] }
    }
    if ($cambios.Count -eq 0) {
        Write-Resultado 'SIN CAMBIOS' "Configuración de '$Titulo'"
        return
    }
    Set-PnPList -Identity $Titulo @cambios | Out-Null
    Write-Resultado 'ACTUALIZADO' ("Configuración de '{0}': {1}" -f $Titulo, (($cambios.Keys | Sort-Object) -join ', '))
}

# Deja una columna con el nombre visible, descripción, obligatoriedad, índice y
# valores únicos del esquema. Devuelve la lista de cambios hechos.
function Sync-Columna {
    param(
        [Parameter(Mandatory)][string]$Lista,
        [Parameter(Mandatory)][string]$NombreInterno,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Especificacion
    )
    $campo = Get-PnPField -List $Lista -Identity $NombreInterno
    $nombreVisible = $NombreInterno
    if ($Especificacion.NombreVisible) { $nombreVisible = [string]$Especificacion.NombreVisible }
    $requerida = [bool]$Especificacion.Requerida
    $unica     = [bool]$Especificacion.ValoresUnicos
    $indexada  = [bool]$Especificacion.Indexada -or $unica   # SharePoint exige índice para valores únicos
    $cambios   = New-Object System.Collections.Generic.List[string]

    # Primero se quitan los valores únicos: sin eso SharePoint no deja quitar el índice.
    if ($campo.EnforceUniqueValues -and -not $unica) {
        Set-PnPField -List $Lista -Identity $NombreInterno -Values @{ EnforceUniqueValues = $false } | Out-Null
        $cambios.Add('sin valores únicos')
    }

    $valores = @{}
    if ([string]$campo.Title -cne $nombreVisible) {
        $valores.Title = $nombreVisible
        $cambios.Add("nombre visible '$nombreVisible'")
    }
    if ($null -ne $Especificacion.Descripcion -and [string]$campo.Description -cne [string]$Especificacion.Descripcion) {
        $valores.Description = [string]$Especificacion.Descripcion
        $cambios.Add('descripción')
    }
    if ([bool]$campo.Required -ne $requerida) {
        $valores.Required = $requerida
        $cambios.Add($(if ($requerida) { 'obligatoria' } else { 'opcional' }))
    }
    if ([bool]$campo.Indexed -ne $indexada) {
        $valores.Indexed = $indexada
        $cambios.Add($(if ($indexada) { 'indexada' } else { 'sin índice' }))
    }
    if ($valores.Count -gt 0) {
        Set-PnPField -List $Lista -Identity $NombreInterno -Values $valores | Out-Null
    }

    # Los valores únicos se activan cuando la columna ya quedó indexada.
    if ($unica -and -not $campo.EnforceUniqueValues) {
        Set-PnPField -List $Lista -Identity $NombreInterno -Values @{ EnforceUniqueValues = $true } | Out-Null
        $cambios.Add('valores únicos')
    }
    return $cambios.ToArray()
}

# Crea la vista o la deja con las columnas, consulta y paginación del esquema.
function Sync-Vista {
    param(
        [Parameter(Mandatory)][string]$Lista,
        [string]$Titulo,
        [Parameter(Mandatory)][string[]]$Campos,
        [Parameter(Mandatory)][string]$Consulta,
        [switch]$Predeterminada
    )
    $vistas = @(Get-PnPView -List $Lista -Includes ViewFields, ViewQuery, RowLimit, Paged)
    if ($Predeterminada) {
        $vista = $vistas | Where-Object { $_.DefaultView } | Select-Object -First 1
        if (-not $vista) { throw "La lista '$Lista' no tiene vista predeterminada." }
        $Titulo = $vista.Title
    }
    else {
        $vista = $vistas | Where-Object { $_.Title -eq $Titulo } | Select-Object -First 1
    }

    if (-not $vista) {
        Add-PnPView -List $Lista -Title $Titulo -Fields $Campos -Query $Consulta -RowLimit 100 -Paged | Out-Null
        Write-Resultado 'CREADO' "Vista '$Titulo'"
        return
    }

    $igual = ((@($vista.ViewFields) -join ',') -ceq ($Campos -join ',')) -and
             ((Format-Caml $vista.ViewQuery) -ceq (Format-Caml $Consulta)) -and
             ($vista.RowLimit -eq 100) -and [bool]$vista.Paged
    if ($igual) {
        Write-Resultado 'SIN CAMBIOS' "Vista '$Titulo'"
        return
    }
    Set-PnPView -List $Lista -Identity $vista.Id -Fields $Campos -Values @{ ViewQuery = $Consulta; RowLimit = [uint32]100; Paged = $true } | Out-Null
    Write-Resultado 'ACTUALIZADO' "Vista '$Titulo': columnas, filtro y orden"
}

# Crea las columnas del esquema que falten en una lista o biblioteca y sincroniza las existentes.
function Sync-ColumnasDeLista {
    param(
        [Parameter(Mandatory)][string]$Lista,
        [AllowEmptyCollection()][object[]]$Columnas = @()
    )
    $camposExistentes = @{}
    foreach ($campo in (Get-PnPField -List $Lista)) { $camposExistentes[$campo.InternalName] = $campo }

    foreach ($columna in @($Columnas)) {
        $nombre = [string]$columna.NombreInterno
        $tipoEsperado = $script:MapaTipos[[string]$columna.Tipo]
        $creada = $false

        if (-not $camposExistentes.ContainsKey($nombre)) {
            Add-PnPFieldFromXml -List $Lista -FieldXml (New-FieldXml -Columna $columna) | Out-Null
            $creada = $true
        }
        elseif ($camposExistentes[$nombre].TypeAsString -ne $tipoEsperado) {
            throw "La columna '$nombre' de '$Lista' ya existe con tipo '$($camposExistentes[$nombre].TypeAsString)' y el esquema pide '$tipoEsperado'. El script no cambia tipos de columna porque se pueden perder datos: corríjala a mano o elimínela si está vacía."
        }

        $cambios = @(Sync-Columna -Lista $Lista -NombreInterno $nombre -Especificacion $columna)
        if ($creada) { Write-Resultado 'CREADO' ("{0}: {1} ({2}) -> {3}" -f $Lista, $nombre, $columna.Tipo, ($cambios -join ', ')) }
        elseif ($cambios.Count -gt 0) { Write-Resultado 'ACTUALIZADO' ("{0}: {1} -> {2}" -f $Lista, $nombre, ($cambios -join ', ')) }
        else { Write-Resultado 'SIN CAMBIOS' ("{0}: {1}" -f $Lista, $nombre) }

        # Las opciones de una columna existente no se cambian solas: el flujo y Power Query dependen de ellas.
        if ($columna.Tipo -eq 'Opcion' -and -not $creada) {
            $definicion = Get-DefinicionCampo -Campo (Get-PnPField -List $Lista -Identity $nombre)
            if (($definicion.Opciones -join '|') -cne (@($columna.Opciones) -join '|') -or $definicion.Predeterminado -cne [string]$columna.ValorPredeterminado) {
                Write-Resultado 'AVISO' ("{0}: en SharePoint las opciones son [{1}] (predeterminada '{2}') y el esquema pide [{3}] (predeterminada '{4}'). Revíselo a mano." -f `
                    $nombre, ($definicion.Opciones -join ', '), $definicion.Predeterminado, (@($columna.Opciones) -join ', '), $columna.ValorPredeterminado)
            }
        }
    }
}

# Agrega a un grupo de SharePoint las personas o grupos de Entra ID que falten.
function Add-MiembrosGrupo {
    param(
        [Parameter(Mandatory)]$Grupo,
        [Parameter(Mandatory)][object[]]$Principales
    )
    $miembros = @(Get-PnPGroupMember -Group $Grupo.Id)
    foreach ($principal in $Principales) {
        $descripcion = '{0} {1}' -f $principal.Tipo, $principal.Valor
        if (Test-EsMiembro -Miembros $miembros -Principal $principal) {
            Write-Resultado 'SIN CAMBIOS' "$descripcion ya está en '$($Grupo.Title)'"
            continue
        }
        try {
            Add-PnPGroupMember -LoginName (Get-LoginParaAgregar -Principal $principal) -Group $Grupo.Id | Out-Null
        }
        catch {
            throw "No se pudo agregar $descripcion al grupo '$($Grupo.Title)': $($_.Exception.Message) Verifique que el correo o el Id. del grupo existan en Entra ID."
        }
        Write-Resultado 'CREADO' "$descripcion agregado a '$($Grupo.Title)'"
    }
}

#endregion

$contexto = Import-ConfigLegalizaciones -RutaConfig $RutaConfig -RutaEsquema $RutaEsquema
$cfg = $contexto.Config
$esq = $contexto.Esquema
$sitioDedicado = $cfg.Permisos.Ambito -eq 'Sitio'

$rutaLog = Start-RegistroEjecucion -Contexto $contexto -Prefijo 'aprovisionamiento'
try {
    # ======================================================================== 1
    Write-Paso "1/8 Sitio $($contexto.Urls.Sitio)"

    $cnxAdmin = $null
    $necesitaAdmin = [bool]$cfg.Sitio.Crear -or ($sitioDedicado -and [bool]$cfg.Sitio.DeshabilitarComparticionExterna)
    if ($necesitaAdmin) {
        Write-Resultado 'INFO' "Conectando al centro de administración $($contexto.Urls.Admin) (requiere rol Administrador de SharePoint)"
        $cnxAdmin = Connect-PnPOnline -Url $contexto.Urls.Admin -Interactive -ClientId $cfg.Autenticacion.ClientId -ReturnConnection
    }

    if ($cfg.Sitio.Crear) {
        $sitioExistente = $null
        try { $sitioExistente = Get-PnPTenantSite -Identity $contexto.Urls.Sitio -Connection $cnxAdmin -ErrorAction Stop }
        catch { $sitioExistente = $null }

        if ($sitioExistente) {
            Write-Resultado 'SIN CAMBIOS' 'El sitio ya existe'
        }
        else {
            Write-Resultado 'INFO' 'Creando el sitio; puede tardar unos minutos...'
            New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $cfg.Sitio.Titulo -Url $contexto.Urls.Sitio -Lcid ([uint32]$cfg.Sitio.Lcid) -Wait -Connection $cnxAdmin | Out-Null
            Write-Resultado 'CREADO' "Sitio '$($cfg.Sitio.Titulo)' (sin Grupo de Microsoft 365)"
        }
    }

    Connect-SitioLegalizaciones -Contexto $contexto
    $web = Get-PnPWeb -Includes Language, Description

    if ($cfg.Sitio.Crear) {
        if ([uint32]$web.Language -ne [uint32]$cfg.Sitio.Lcid) {
            Write-Resultado 'AVISO' "El sitio tiene idioma $($web.Language) y la configuración pide $($cfg.Sitio.Lcid). El idioma de un sitio no se puede cambiar después de crearlo."
        }
        if ([string]$web.Description -cne [string]$cfg.Sitio.Descripcion) {
            Set-PnPWeb -Description $cfg.Sitio.Descripcion | Out-Null
            Write-Resultado 'ACTUALIZADO' 'Descripción del sitio'
        }
    }

    # ======================================================================== 2
    Write-Paso '2/8 Configuración regional'

    $regional = Get-PnPProperty -ClientObject $web -Property RegionalSettings
    Get-PnPProperty -ClientObject $regional -Property LocaleId, TimeZone, TimeZones | Out-Null

    if (-not $sitioDedicado) {
        Write-Resultado 'AVISO' "Sitio compartido: no se cambia su configuración regional. Actual: $($regional.TimeZone.Description); LocaleId $($regional.LocaleId)."
    }
    else {
        $zonas = @($regional.TimeZones | Where-Object { $_.Description -like "*$($contexto.Region.ZonaSharePoint)*" })
        if ($zonas.Count -ne 1) {
            $disponibles = ($regional.TimeZones | ForEach-Object { '    ' + $_.Description }) -join "`n"
            throw "La zona horaria '$($contexto.Region.ZonaSharePoint)' (Pais = $($contexto.Region.Pais)) coincide con $($zonas.Count) zonas horarias y debe coincidir con una sola. Zonas disponibles:`n$disponibles"
        }
        $zona = $zonas[0]
        $cambios = @()
        if ($regional.TimeZone.Id -ne $zona.Id) {
            $regional.TimeZone = $zona
            $cambios += "zona horaria $($zona.Description)"
        }
        if ([uint32]$regional.LocaleId -ne [uint32]$contexto.Region.LocaleId) {
            $regional.LocaleId = [uint32]$contexto.Region.LocaleId
            $cambios += "configuración regional $($contexto.Region.LocaleId)"
        }
        if ($cambios.Count -gt 0) {
            $web.Update()
            Invoke-PnPQuery
            Write-Resultado 'ACTUALIZADO' ($cambios -join '; ')
        }
        else {
            Write-Resultado 'SIN CAMBIOS' "Zona horaria $($zona.Description); configuración regional $($regional.LocaleId)"
        }
    }

    # ======================================================================== 3
    Write-Paso "3/8 Biblioteca $($esq.Biblioteca.Titulo)"

    $bib = $esq.Biblioteca
    $biblioteca = Get-PnPList -Identity $bib.Titulo -Includes BaseTemplate
    if (-not $biblioteca) {
        New-PnPList -Title $bib.Titulo -Url $bib.Url -Template DocumentLibrary -OnQuickLaunch | Out-Null
        Write-Resultado 'CREADO' "Biblioteca '$($bib.Titulo)'"
    }
    elseif ($biblioteca.BaseTemplate -ne 101) {
        throw "Ya existe una lista llamada '$($bib.Titulo)' que no es una biblioteca de documentos. Cámbiele el nombre o cambie Biblioteca.Titulo en el esquema."
    }
    else {
        Write-Resultado 'SIN CAMBIOS' "La biblioteca '$($bib.Titulo)' ya existe"
    }

    # Sin desprotección obligatoria: Power Automate sube los archivos y deben quedar publicados.
    Sync-PropiedadesLista -Titulo $bib.Titulo -Deseado ([ordered]@{
        Description         = $bib.Descripcion
        EnableVersioning    = $true
        EnableMinorVersions = $false
        ForceCheckout       = $false
    })

    foreach ($carpeta in $bib.CarpetasRaiz) {
        $existe = Get-PnPFolderItem -FolderSiteRelativeUrl $bib.Url -ItemType Folder -ItemName $carpeta
        if ($existe) {
            Write-Resultado 'SIN CAMBIOS' "Carpeta $($bib.Url)/$carpeta"
        }
        else {
            Resolve-PnPFolder -SiteRelativePath "$($bib.Url)/$carpeta" | Out-Null
            Write-Resultado 'CREADO' "Carpeta $($bib.Url)/$carpeta"
        }
    }

    # Columna ID Legalización de los archivos y vista predeterminada de la biblioteca.
    Sync-ColumnasDeLista -Lista $bib.Titulo -Columnas @($bib.Columnas)
    if ($bib.VistaPredeterminada) {
        Sync-Vista -Lista $bib.Titulo -Campos $bib.VistaPredeterminada.Campos -Consulta $bib.VistaPredeterminada.Consulta -Predeterminada
    }

    # ======================================================================== 4
    Write-Paso "4/8 Lista $($esq.Lista.Titulo)"

    $lis = $esq.Lista
    $lista = Get-PnPList -Identity $lis.Titulo -Includes BaseTemplate
    if (-not $lista) {
        New-PnPList -Title $lis.Titulo -Url $lis.Url -Template GenericList -OnQuickLaunch | Out-Null
        Write-Resultado 'CREADO' "Lista '$($lis.Titulo)'"
    }
    elseif ($lista.BaseTemplate -ne 100) {
        throw "Ya existe una lista llamada '$($lis.Titulo)' que no es una lista personalizada. Cámbiele el nombre o cambie Lista.Titulo en el esquema."
    }
    else {
        Write-Resultado 'SIN CAMBIOS' "La lista '$($lis.Titulo)' ya existe"
    }

    # Historial de versiones = trazabilidad de cada cambio de estado (quién y cuándo).
    # Sin adjuntos: los soportes viven solo en la biblioteca.
    Sync-PropiedadesLista -Titulo $lis.Titulo -Deseado ([ordered]@{
        Description          = $lis.Descripcion
        EnableVersioning     = $true
        MajorVersions        = [uint32]$lis.VersionesPrincipales
        EnableAttachments    = $false
        EnableFolderCreation = $false
    })

    # ======================================================================== 5
    Write-Paso '5/8 Columnas del registro'

    $cambios = @(Sync-Columna -Lista $lis.Titulo -NombreInterno 'Title' -Especificacion $lis.ColumnaTitulo)
    if ($cambios.Count -gt 0) { Write-Resultado 'ACTUALIZADO' ("Title -> {0}" -f ($cambios -join ', ')) }
    else { Write-Resultado 'SIN CAMBIOS' "Title ('$($lis.ColumnaTitulo.NombreVisible)')" }

    Sync-ColumnasDeLista -Lista $lis.Titulo -Columnas @($lis.Columnas)

    # ======================================================================== 6
    Write-Paso '6/8 Vistas'

    Sync-Vista -Lista $lis.Titulo -Campos $lis.VistaPredeterminada.Campos -Consulta $lis.VistaPredeterminada.Consulta -Predeterminada
    foreach ($vista in $lis.Vistas) {
        Sync-Vista -Lista $lis.Titulo -Titulo $vista.Titulo -Campos $vista.Campos -Consulta $vista.Consulta
    }

    # ======================================================================== 7
    Write-Paso '7/8 Permisos'

    # Los nombres de los niveles dependen del idioma del sitio; se buscan por tipo.
    $rolControlTotal = Get-NombreRol -Tipo Administrator
    $rolColaborar    = Get-NombreRol -Tipo Contributor
    $rolEdicion      = Get-NombreRol -Tipo Editor

    $principalesConta = @($cfg.Permisos.Contabilidad)
    if ($cfg.Permisos.CuentaServicioFlujo) {
        $principalesConta += @{ Tipo = 'Usuario'; Valor = [string]$cfg.Permisos.CuentaServicioFlujo }
    }
    else {
        Write-Resultado 'AVISO' 'Permisos.CuentaServicioFlujo está vacío: la cuenta de Power Automate no tendrá acceso hasta que la configure y vuelva a ejecutar este script.'
    }

    $grupoPropietarios = Get-PnPGroup -AssociatedOwnerGroup
    if (-not $grupoPropietarios) { throw 'El sitio no tiene un grupo de propietarios asociado.' }
    Add-MiembrosGrupo -Grupo $grupoPropietarios -Principales @($cfg.Permisos.PropietariosTI)

    if ($sitioDedicado) {
        # Contabilidad entra al grupo de Miembros del sitio con nivel Colaborar en lugar de Edición.
        # Edición permitiría borrar la lista o cambiar sus columnas.
        $grupoMiembros = Get-PnPGroup -AssociatedMemberGroup
        if (-not $grupoMiembros) { throw "El sitio no tiene un grupo de miembros asociado. Use Permisos.Ambito = 'ListaYBiblioteca'." }

        $roles = @(Get-RolesGrupo -Grupo $grupoMiembros)
        if ($roles -notcontains 'Contributor') {
            Set-PnPGroupPermissions -Identity $grupoMiembros.Id -AddRole $rolColaborar | Out-Null
            Write-Resultado 'ACTUALIZADO' "Grupo '$($grupoMiembros.Title)': nivel '$rolColaborar' asignado"
        }
        if ($roles -contains 'Editor') {
            Set-PnPGroupPermissions -Identity $grupoMiembros.Id -RemoveRole $rolEdicion | Out-Null
            Write-Resultado 'ACTUALIZADO' "Grupo '$($grupoMiembros.Title)': nivel '$rolEdicion' retirado"
        }
        if (($roles -contains 'Contributor') -and ($roles -notcontains 'Editor')) {
            Write-Resultado 'SIN CAMBIOS' "Grupo '$($grupoMiembros.Title)' con nivel '$rolColaborar'"
        }
        Add-MiembrosGrupo -Grupo $grupoMiembros -Principales $principalesConta
    }
    else {
        # Sitio compartido: la lista y la biblioteca dejan de heredar y solo acceden
        # los propietarios del sitio y el grupo de Contabilidad.
        $grupoConta = Get-PnPGroup | Where-Object { $_.Title -eq $cfg.Permisos.NombreGrupoSharePoint } | Select-Object -First 1
        if (-not $grupoConta) {
            $grupoConta = New-PnPGroup -Title $cfg.Permisos.NombreGrupoSharePoint -Description 'Contabilidad: revisión de legalizaciones de tarjeta corporativa.' -Owner $grupoPropietarios.Title
            Write-Resultado 'CREADO' "Grupo de SharePoint '$($cfg.Permisos.NombreGrupoSharePoint)'"
        }
        else {
            Write-Resultado 'SIN CAMBIOS' "Grupo de SharePoint '$($cfg.Permisos.NombreGrupoSharePoint)'"
        }
        Add-MiembrosGrupo -Grupo $grupoConta -Principales $principalesConta

        foreach ($titulo in @($esq.Biblioteca.Titulo, $esq.Lista.Titulo)) {
            $objeto = Get-PnPList -Identity $titulo -Includes HasUniqueRoleAssignments
            if (-not $objeto.HasUniqueRoleAssignments) {
                Set-PnPList -Identity $titulo -BreakRoleInheritance -ClearSubScopes | Out-Null
                Write-Resultado 'ACTUALIZADO' "'$titulo' deja de heredar los permisos del sitio"
            }
            Set-PnPListPermission -Identity $titulo -Group $grupoPropietarios.Title -AddRole $rolControlTotal | Out-Null
            Set-PnPListPermission -Identity $titulo -Group $grupoConta.Title -AddRole $rolColaborar | Out-Null
            Write-Resultado 'INFO' "'$titulo': '$($grupoPropietarios.Title)' = $rolControlTotal; '$($grupoConta.Title)' = $rolColaborar"
        }
    }

    # ======================================================================== 8
    Write-Paso '8/8 Uso compartido'

    if (-not $sitioDedicado) {
        Write-Resultado 'AVISO' 'Sitio compartido: no se cambian sus opciones de uso compartido porque afectan a otras áreas. Pida a TI que las revise.'
    }
    else {
        if ($cfg.Sitio.DeshabilitarComparticionExterna) {
            $sitioAdmin = Get-PnPTenantSite -Identity $contexto.Urls.Sitio -Connection $cnxAdmin
            if ([string]$sitioAdmin.SharingCapability -eq 'Disabled') {
                Write-Resultado 'SIN CAMBIOS' 'Uso compartido externo deshabilitado'
            }
            else {
                Set-PnPTenantSite -Identity $contexto.Urls.Sitio -SharingCapability Disabled -Connection $cnxAdmin | Out-Null
                Write-Resultado 'ACTUALIZADO' 'Uso compartido externo deshabilitado'
            }
        }
        if ($cfg.Sitio.SoloPropietariosComparten) {
            $webCompartir = Get-PnPWeb -Includes MembersCanShare
            if (-not $webCompartir.MembersCanShare) {
                Write-Resultado 'SIN CAMBIOS' 'Solo los propietarios pueden compartir'
            }
            else {
                Set-PnPWeb -MembersCanShare:$false | Out-Null
                Write-Resultado 'ACTUALIZADO' 'Solo los propietarios pueden compartir'
            }
        }
    }

    # ======================================================================== Resumen
    Write-Paso 'Resumen'
    foreach ($estado in $script:Contadores.Keys) {
        Write-Host ('  {0,-12} {1}' -f $estado, $script:Contadores[$estado])
    }
    Write-Host ''
    Write-Host "  Sitio:      $($contexto.Urls.Sitio)"
    Write-Host "  Biblioteca: $($contexto.Urls.Sitio)/$($esq.Biblioteca.Url)"
    Write-Host "  Registro:   $($contexto.Urls.Sitio)/$($esq.Lista.Url)"
    Write-Host "  Log:        $rutaLog"
    Write-Host ''
    Write-Host '  Siguiente paso: ./src/4-infraestructura/sharepoint/02-Verificar-SharePoint.ps1' -ForegroundColor Green
}
catch {
    Write-Resultado 'ERROR' $_.Exception.Message
    Write-Host "  El detalle quedó en: $rutaLog" -ForegroundColor Red
    throw
}
finally {
    Stop-Transcript | Out-Null
}
