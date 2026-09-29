#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '3.0.0' }
<#
.SYNOPSIS
    Verifica, SIN hacer cambios, que SharePoint quedó como indican la configuración y el esquema.

.DESCRIPTION
    Revisa el sitio, la biblioteca y sus carpetas, la lista, cada columna (tipo, nombre
    visible, obligatoriedad, índice, opciones), las vistas y los permisos.
    Muestra una tabla con OK / AVISO / FALLA, la guarda en CSV junto al log y termina
    con código de salida 1 si hay alguna FALLA.

.PARAMETER RutaConfig
    Configuración del entorno. Por defecto, Config.Legalizaciones.psd1.

.PARAMETER RutaEsquema
    Contrato de datos. Por defecto, Esquema.Legalizaciones.psd1.

.EXAMPLE
    ./02-Verificar-SharePoint.ps1
#>
[CmdletBinding()]
param(
    [string]$RutaConfig  = (Join-Path $PSScriptRoot '..\..\..\config\Config.Legalizaciones.psd1'),
    [string]$RutaEsquema = (Join-Path $PSScriptRoot '..\..\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Comun.Legalizaciones.ps1')

#region Funciones de este script -------------------------------------------------

$script:Resultados = New-Object System.Collections.Generic.List[object]

# Registra una verificación. Si no se indica -Cumple, compara Esperado y Actual como texto.
function Add-Resultado {
    param(
        [Parameter(Mandatory)][string]$Area,
        [Parameter(Mandatory)][string]$Verificacion,
        [AllowNull()]$Esperado,
        [AllowNull()]$Actual,
        [Nullable[bool]]$Cumple = $null,
        [switch]$SoloAviso
    )
    if ($null -eq $Cumple) { $Cumple = ([string]$Esperado) -ceq ([string]$Actual) }
    $estado = if ($Cumple) { 'OK' } elseif ($SoloAviso) { 'AVISO' } else { 'FALLA' }
    $script:Resultados.Add([pscustomobject]@{
        Area         = $Area
        Verificacion = $Verificacion
        Esperado     = [string]$Esperado
        Actual       = [string]$Actual
        Resultado    = $estado
    })
}

function ConvertTo-SiNo {
    param([AllowNull()]$Valor)
    if ($Valor) { return 'Sí' }
    return 'No'
}

# Permisos asignados directamente en un sitio, lista o biblioteca.
function Get-AsignacionesRol {
    param([Parameter(Mandatory)]$Objeto)
    $salida = @()
    foreach ($asignacion in (Get-PnPProperty -ClientObject $Objeto -Property RoleAssignments)) {
        Get-PnPProperty -ClientObject $asignacion -Property Member, RoleDefinitionBindings | Out-Null
        $salida += [pscustomobject]@{
            Principal = $asignacion.Member.Title
            LoginName = $asignacion.Member.LoginName
            Permisos  = (@($asignacion.RoleDefinitionBindings) | ForEach-Object { $_.Name }) -join ', '
        }
    }
    return $salida
}

# "Todos" y "Todos excepto los usuarios externos".
$patronAccesoGeneral = 'spo-grid-all-users|^c:0\(\.s\|true$'

# Revisa que en un objeto solo tengan permisos los principales esperados.
function Test-AsignacionesEsperadas {
    param(
        [Parameter(Mandatory)][string]$Area,
        [Parameter(Mandatory)][string]$Donde,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Asignaciones,
        [Parameter(Mandatory)][string[]]$PrincipalesEsperados
    )
    $generales = @($Asignaciones | Where-Object { $_.LoginName -match $patronAccesoGeneral })
    $actual = 'Ninguno'
    if ($generales.Count -gt 0) { $actual = ($generales | ForEach-Object { $_.Principal }) -join ', ' }
    Add-Resultado $Area "$Donde sin acceso para 'Todos' ni 'Todos excepto los usuarios externos'" 'Ninguno' $actual

    $otros = @($Asignaciones | Where-Object { $PrincipalesEsperados -notcontains $_.Principal -and $_.LoginName -notmatch $patronAccesoGeneral })
    $actual = 'Ninguno'
    if ($otros.Count -gt 0) { $actual = ($otros | ForEach-Object { '{0} ({1})' -f $_.Principal, $_.Permisos }) -join '; ' }
    Add-Resultado $Area "$Donde sin permisos adicionales a los esperados" 'Ninguno' $actual -SoloAviso

    Write-Host "  Permisos en $($Donde):" -ForegroundColor Cyan
    $Asignaciones | Format-Table Principal, Permisos, LoginName -AutoSize | Out-Host
}

# Revisa que cada principal de la configuración esté en el grupo.
function Test-MiembrosGrupo {
    param(
        [Parameter(Mandatory)]$Grupo,
        [Parameter(Mandatory)][object[]]$Principales
    )
    $miembros = @(Get-PnPGroupMember -Group $Grupo.Id)
    foreach ($principal in $Principales) {
        $esMiembro = Test-EsMiembro -Miembros $miembros -Principal $principal
        Add-Resultado 'Permisos' ("{0} {1} en '{2}'" -f $principal.Tipo, $principal.Valor, $Grupo.Title) 'Sí' (ConvertTo-SiNo $esMiembro)
    }
}

#endregion

$contexto = Import-ConfigLegalizaciones -RutaConfig $RutaConfig -RutaEsquema $RutaEsquema
$cfg = $contexto.Config
$esq = $contexto.Esquema
$sitioDedicado = $cfg.Permisos.Ambito -eq 'Sitio'
$fallas = 0

$rutaLog = Start-RegistroEjecucion -Contexto $contexto -Prefijo 'verificacion'
try {
    Connect-SitioLegalizaciones -Contexto $contexto

    # ----------------------------------------------------------------- Sitio
    Write-Paso 'Verificando el sitio'
    $web = Get-PnPWeb -Includes Language
    Add-Resultado 'Sitio' 'URL' $contexto.Urls.Sitio.ToLowerInvariant() ([string]$web.Url).ToLowerInvariant()
    Add-Resultado 'Sitio' 'Idioma (LCID)' $cfg.Sitio.Lcid $web.Language -SoloAviso:(-not $cfg.Sitio.Crear)

    $regional = Get-PnPProperty -ClientObject $web -Property RegionalSettings
    Get-PnPProperty -ClientObject $regional -Property LocaleId, TimeZone | Out-Null
    $zonaActual = [string]$regional.TimeZone.Description
    Add-Resultado 'Sitio' 'Zona horaria' "contiene '$($contexto.Region.ZonaSharePoint)'" $zonaActual -Cumple ($zonaActual -like "*$($contexto.Region.ZonaSharePoint)*") -SoloAviso:(-not $sitioDedicado)
    Add-Resultado 'Sitio' 'Configuración regional (LocaleId)' $contexto.Region.LocaleId $regional.LocaleId -SoloAviso:(-not $sitioDedicado)

    if ($sitioDedicado -and $cfg.Sitio.SoloPropietariosComparten) {
        try {
            $webCompartir = Get-PnPWeb -Includes MembersCanShare
            Add-Resultado 'Sitio' 'Solo los propietarios pueden compartir (MembersCanShare)' 'False' $webCompartir.MembersCanShare
        }
        catch {
            Add-Resultado 'Sitio' 'Solo los propietarios pueden compartir (MembersCanShare)' 'False' 'No se pudo leer' -SoloAviso
        }
    }

    # ----------------------------------------------------------- Biblioteca
    Write-Paso 'Verificando la biblioteca'
    $bib = $esq.Biblioteca
    $biblioteca = Get-PnPList -Identity $bib.Titulo -Includes BaseTemplate, EnableVersioning, ForceCheckout, HasUniqueRoleAssignments
    Add-Resultado 'Biblioteca' "Existe '$($bib.Titulo)'" 'Sí' (ConvertTo-SiNo $biblioteca)
    if ($biblioteca) {
        Add-Resultado 'Biblioteca' 'Es biblioteca de documentos (101)' 101 $biblioteca.BaseTemplate
        Add-Resultado 'Biblioteca' 'Historial de versiones' 'True' $biblioteca.EnableVersioning
        Add-Resultado 'Biblioteca' 'Desprotección obligatoria' 'False' $biblioteca.ForceCheckout
        $carpetas = @(Get-PnPFolderItem -FolderSiteRelativeUrl $bib.Url -ItemType Folder | ForEach-Object { $_.Name })
        foreach ($carpeta in $bib.CarpetasRaiz) {
            Add-Resultado 'Biblioteca' "Carpeta '$carpeta'" 'Sí' (ConvertTo-SiNo ($carpetas -contains $carpeta))
        }
        $camposBiblioteca = @{}
        foreach ($campo in (Get-PnPField -List $bib.Titulo)) { $camposBiblioteca[$campo.InternalName] = $campo }
        foreach ($columna in @($bib.Columnas)) {
            $nombre = [string]$columna.NombreInterno
            $campo = $camposBiblioteca[$nombre]
            Add-Resultado 'Biblioteca' "Columna $nombre existe" 'Sí' (ConvertTo-SiNo $campo)
            if ($campo) {
                Add-Resultado 'Biblioteca' "$nombre tipo" $script:MapaTipos[[string]$columna.Tipo] $campo.TypeAsString
                Add-Resultado 'Biblioteca' "$nombre indexada" ([bool]$columna.Indexada) ([bool]$campo.Indexed)
            }
        }
    }

    # ---------------------------------------------------------------- Lista
    Write-Paso 'Verificando la lista del registro'
    $lis = $esq.Lista
    $lista = Get-PnPList -Identity $lis.Titulo -Includes BaseTemplate, EnableVersioning, EnableAttachments, HasUniqueRoleAssignments
    Add-Resultado 'Lista' "Existe '$($lis.Titulo)'" 'Sí' (ConvertTo-SiNo $lista)
    if ($lista) {
        Add-Resultado 'Lista' 'Es lista personalizada (100)' 100 $lista.BaseTemplate
        Add-Resultado 'Lista' 'Historial de versiones' 'True' $lista.EnableVersioning
        Add-Resultado 'Lista' 'Adjuntos deshabilitados' 'False' $lista.EnableAttachments

        # ------------------------------------------------------------ Columnas
        Write-Paso 'Verificando columnas'
        $campos = @{}
        foreach ($campo in (Get-PnPField -List $lis.Titulo)) { $campos[$campo.InternalName] = $campo }

        $titulo = $lis.ColumnaTitulo
        $especificaciones = @(@{
                NombreInterno = 'Title'
                NombreVisible = $titulo.NombreVisible
                Tipo          = 'Texto'
                Requerida     = $titulo.Requerida
                Indexada      = $titulo.Indexada
                ValoresUnicos = $titulo.ValoresUnicos
            }) + @($lis.Columnas)

        foreach ($columna in $especificaciones) {
            $nombre = [string]$columna.NombreInterno
            $campo = $campos[$nombre]
            Add-Resultado 'Columnas' "$nombre existe" 'Sí' (ConvertTo-SiNo $campo)
            if (-not $campo) { continue }

            $visible = $nombre
            if ($columna.NombreVisible) { $visible = [string]$columna.NombreVisible }
            $unica = [bool]$columna.ValoresUnicos
            Add-Resultado 'Columnas' "$nombre tipo" $script:MapaTipos[[string]$columna.Tipo] $campo.TypeAsString
            Add-Resultado 'Columnas' "$nombre nombre visible" $visible $campo.Title
            Add-Resultado 'Columnas' "$nombre obligatoria" ([bool]$columna.Requerida) ([bool]$campo.Required)
            Add-Resultado 'Columnas' "$nombre indexada" ([bool]$columna.Indexada -or $unica) ([bool]$campo.Indexed)
            if ($unica) { Add-Resultado 'Columnas' "$nombre valores únicos" $true ([bool]$campo.EnforceUniqueValues) }

            $definicion = Get-DefinicionCampo -Campo $campo
            switch ([string]$columna.Tipo) {
                'FechaHora' { Add-Resultado 'Columnas' "$nombre formato" 'DateTime' $definicion.Formato }
                'Fecha'     { Add-Resultado 'Columnas' "$nombre formato" 'DateOnly' $definicion.Formato }
                'Opcion' {
                    Add-Resultado 'Columnas' "$nombre opciones" (@($columna.Opciones) -join ' | ') ($definicion.Opciones -join ' | ')
                    Add-Resultado 'Columnas' "$nombre valor predeterminado" $columna.ValorPredeterminado $definicion.Predeterminado
                }
            }
        }

        # --------------------------------------------------------------- Vistas
        Write-Paso 'Verificando vistas'
        $vistas = @(Get-PnPView -List $lis.Titulo -Includes ViewQuery)
        foreach ($vista in $lis.Vistas) {
            Add-Resultado 'Vistas' "Vista '$($vista.Titulo)'" 'Sí' (ConvertTo-SiNo ($vistas | Where-Object { $_.Title -eq $vista.Titulo }))
        }
        $predeterminada = $vistas | Where-Object { $_.DefaultView } | Select-Object -First 1
        $sinFiltro = $predeterminada -and ([string]$predeterminada.ViewQuery -notmatch '<Where')
        Add-Resultado 'Vistas' 'Vista predeterminada sin filtro (la puede usar Power Query)' 'Sí' (ConvertTo-SiNo $sinFiltro)
    }

    # ------------------------------------------------------------- Permisos
    Write-Paso 'Verificando permisos'
    $principalesConta = @($cfg.Permisos.Contabilidad)
    if ($cfg.Permisos.CuentaServicioFlujo) {
        $principalesConta += @{ Tipo = 'Usuario'; Valor = [string]$cfg.Permisos.CuentaServicioFlujo }
    }
    else {
        Add-Resultado 'Permisos' 'Cuenta de servicio del flujo configurada' 'Sí' 'No' -SoloAviso
    }

    $grupoPropietarios = Get-PnPGroup -AssociatedOwnerGroup
    Add-Resultado 'Permisos' 'El sitio tiene grupo de propietarios asociado' 'Sí' (ConvertTo-SiNo $grupoPropietarios)
    if ($grupoPropietarios) { Test-MiembrosGrupo -Grupo $grupoPropietarios -Principales @($cfg.Permisos.PropietariosTI) }

    if ($sitioDedicado) {
        $grupoMiembros  = Get-PnPGroup -AssociatedMemberGroup
        $grupoVisitante = Get-PnPGroup -AssociatedVisitorGroup
        Add-Resultado 'Permisos' 'El sitio tiene grupo de miembros asociado' 'Sí' (ConvertTo-SiNo $grupoMiembros)
        if ($grupoMiembros) {
            $roles = @(Get-RolesGrupo -Grupo $grupoMiembros)
            Add-Resultado 'Permisos' "Grupo '$($grupoMiembros.Title)' con nivel Colaborar" 'Sí' (ConvertTo-SiNo ($roles -contains 'Contributor'))
            Add-Resultado 'Permisos' "Grupo '$($grupoMiembros.Title)' sin nivel Edición" 'Sí' (ConvertTo-SiNo ($roles -notcontains 'Editor'))
            Test-MiembrosGrupo -Grupo $grupoMiembros -Principales $principalesConta
        }

        if ($grupoVisitante) {
            $visitantes = @(Get-PnPGroupMember -Group $grupoVisitante.Id | ForEach-Object { $_.Title })
            $actual = 'Ninguno'
            if ($visitantes.Count -gt 0) { $actual = $visitantes -join ', ' }
            Add-Resultado 'Permisos' "Grupo '$($grupoVisitante.Title)' vacío (tendrían lectura)" 'Ninguno' $actual -SoloAviso
        }
        if ($biblioteca) { Add-Resultado 'Permisos' 'La biblioteca hereda los permisos del sitio' 'False' $biblioteca.HasUniqueRoleAssignments -SoloAviso }
        if ($lista) { Add-Resultado 'Permisos' 'La lista hereda los permisos del sitio' 'False' $lista.HasUniqueRoleAssignments -SoloAviso }

        $esperados = @($grupoPropietarios.Title, $grupoMiembros.Title)
        if ($grupoVisitante) { $esperados += $grupoVisitante.Title }
        Test-AsignacionesEsperadas -Area 'Permisos' -Donde 'el sitio' -Asignaciones @(Get-AsignacionesRol -Objeto $web) -PrincipalesEsperados $esperados
    }
    else {
        $grupoConta = Get-PnPGroup | Where-Object { $_.Title -eq $cfg.Permisos.NombreGrupoSharePoint } | Select-Object -First 1
        Add-Resultado 'Permisos' "Existe el grupo '$($cfg.Permisos.NombreGrupoSharePoint)'" 'Sí' (ConvertTo-SiNo $grupoConta)
        if ($grupoConta) { Test-MiembrosGrupo -Grupo $grupoConta -Principales $principalesConta }

        $esperados = @($grupoPropietarios.Title, $cfg.Permisos.NombreGrupoSharePoint)
        foreach ($objeto in @($biblioteca, $lista)) {
            if (-not $objeto) { continue }
            Add-Resultado 'Permisos' "'$($objeto.Title)' con permisos propios (no hereda del sitio)" 'True' $objeto.HasUniqueRoleAssignments
            Test-AsignacionesEsperadas -Area 'Permisos' -Donde "'$($objeto.Title)'" -Asignaciones @(Get-AsignacionesRol -Objeto $objeto) -PrincipalesEsperados $esperados
        }
    }

    # ------------------------------------------------------------- Resultado
    Write-Paso 'Resultado'
    $script:Resultados | Format-Table Area, Verificacion, Esperado, Actual, Resultado -AutoSize -Wrap | Out-Host

    $rutaCsv = [System.IO.Path]::ChangeExtension($rutaLog, '.csv')
    $codificacion = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $script:Resultados | Export-Csv -LiteralPath $rutaCsv -NoTypeInformation -Encoding $codificacion

    $fallas = @($script:Resultados | Where-Object { $_.Resultado -eq 'FALLA' }).Count
    $avisos = @($script:Resultados | Where-Object { $_.Resultado -eq 'AVISO' }).Count
    $correctas = @($script:Resultados | Where-Object { $_.Resultado -eq 'OK' }).Count
    Write-Host ("  OK: {0}   AVISO: {1}   FALLA: {2}" -f $correctas, $avisos, $fallas)
    Write-Host "  Resultados en CSV: $rutaCsv"
    if ($fallas -gt 0) {
        Write-Host '  Hay FALLAS: vuelva a ejecutar 01-Aprovisionar-SharePoint.ps1 y revise los mensajes.' -ForegroundColor Red
    }
    else {
        Write-Host '  SharePoint quedó como indica la configuración.' -ForegroundColor Green
    }
}
catch {
    Write-Resultado 'ERROR' $_.Exception.Message
    Write-Host "  El detalle quedó en: $rutaLog" -ForegroundColor Red
    throw
}
finally {
    Stop-Transcript | Out-Null
}

if ($fallas -gt 0) { exit 1 }
