<#
    Legalizaciones.Dominio.psm1 · Capa de DOMINIO

    Implementación de referencia de las reglas de negocio definidas en dominio.json.
    - No accede a SharePoint, Outlook ni Excel: recibe datos y devuelve resultados (funciones puras).
    - En producción estas mismas reglas las ejecuta Power Automate con expresiones
      (ver src/3-adaptadores/power-automate/expresiones.md). Cada función indica la
      expresión equivalente, y las pruebas (pruebas/unitarias/Dominio.Tests.ps1) fijan
      con ejemplos el comportamiento que el flujo debe reproducir.
    - Compatible con Windows PowerShell 5.1 y PowerShell 7.
#>

Set-StrictMode -Version 2.0

$script:Reglas = $null

#region Reglas ------------------------------------------------------------------

<#
.SYNOPSIS
    Carga las reglas de negocio desde dominio.json (por defecto, el que está junto a este módulo).
#>
function Import-ReglasDominio {
    [CmdletBinding()]
    param([string]$Ruta = (Join-Path $PSScriptRoot 'dominio.json'))
    $texto = [System.IO.File]::ReadAllText($Ruta, [System.Text.Encoding]::UTF8)
    $script:Reglas = $texto | ConvertFrom-Json
    return $script:Reglas
}

<#
.SYNOPSIS
    Devuelve las reglas cargadas (las carga la primera vez).
#>
function Get-ReglasDominio {
    [CmdletBinding()]
    param()
    if ($null -eq $script:Reglas) { [void](Import-ReglasDominio) }
    return $script:Reglas
}

# Tipos de documento en el orden del dominio (factura, comprobante, xml).
function Get-TiposDocumento {
    [CmdletBinding()]
    param()
    $reglas = Get-ReglasDominio
    foreach ($propiedad in $reglas.documentos.PSObject.Properties) {
        [pscustomobject]@{
            Clave             = $propiedad.Name
            Nombre            = $propiedad.Value.nombre
            Obligatorio       = [bool]$propiedad.Value.obligatorio
            Extensiones       = @($propiedad.Value.extensiones)
            ComparaDuplicados = [bool]$propiedad.Value.comparaDuplicados
        }
    }
}

<#
.SYNOPSIS
    Todas las extensiones aceptadas, en el orden del dominio (".pdf", ".xlsx", ".xml").
#>
function Get-ExtensionesPermitidas {
    [CmdletBinding()]
    param()
    $extensiones = @()
    foreach ($tipo in Get-TiposDocumento) { $extensiones += $tipo.Extensiones }
    return $extensiones
}

#endregion

#region Correo ------------------------------------------------------------------

<#
.SYNOPSIS
    Normaliza un asunto como lo hace la condición del desencadenador del flujo.
.DESCRIPTION
    Minúsculas; á é í ó ú -> a e i o u; dos espacios seguidos -> uno (dos pasadas).
    Expresión equivalente (Power Automate):
      replace(replace(replace(replace(replace(replace(replace(toLower(coalesce(triggerBody()?['subject'], '')),
        'á', 'a'), 'é', 'e'), 'í', 'i'), 'ó', 'o'), 'ú', 'u'), '  ', ' '), '  ', ' ')
#>
function ConvertTo-AsuntoNormalizado {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Asunto)
    $texto = ([string]$Asunto).ToLowerInvariant()
    foreach ($par in @(@('á', 'a'), @('é', 'e'), @('í', 'i'), @('ó', 'o'), @('ú', 'u'))) {
        $texto = $texto.Replace($par[0], $par[1])
    }
    return $texto.Replace('  ', ' ').Replace('  ', ' ')
}

<#
.SYNOPSIS
    ¿El correo es una legalización? (condición del desencadenador de LEG-01).
#>
function Test-AsuntoLegalizacion {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Asunto)
    $clave = (Get-ReglasDominio).correo.textoClaveAsunto
    return (ConvertTo-AsuntoNormalizado $Asunto).Contains($clave)
}

<#
.SYNOPSIS
    Nombre del colaborador escrito en el asunto: el texto entre el primer " - " y el siguiente " - ".
    Devuelve '' si el asunto no trae el separador.
.DESCRIPTION
    Expresión equivalente (LEG-01, acción Nombre_en_asunto; E-02 del catálogo):
      trim(coalesce(split(outputs('Asunto'), ' - ')?[1], ''))
    donde outputs('Asunto') = trim(coalesce(triggerBody()?['subject'], '')).
    Se usa split + ?[1] en lugar de indexOf/substring porque Power Automate evalúa las dos
    ramas de if(): substring fallaría con asuntos cortos sin separador.
#>
function Get-NombreEnAsunto {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Asunto)
    $separador = (Get-ReglasDominio).correo.separadorNombre
    $texto = ([string]$Asunto).Trim()
    $posicion = $texto.IndexOf($separador, [System.StringComparison]::Ordinal)
    if ($posicion -lt 0) { return '' }
    $resto = $texto.Substring($posicion + $separador.Length)
    return $resto.Split([string[]]@($separador), [System.StringSplitOptions]::None)[0].Trim()
}

<#
.SYNOPSIS
    Colaborador del caso: el nombre del asunto o, si no viene, la parte del correo antes de la @.
.DESCRIPTION
    Expresión equivalente (acción Colaborador):
      if(empty(outputs('Nombre_en_asunto')), first(split(outputs('Remitente'), '@')), outputs('Nombre_en_asunto'))
    donde outputs('Remitente') = toLower(trim(coalesce(triggerBody()?['from'], ''))).
#>
function Get-NombreColaborador {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$Asunto,
        [AllowNull()][AllowEmptyString()][string]$Remitente
    )
    $nombre = Get-NombreEnAsunto $Asunto
    if ($nombre) { return $nombre }
    return ([string]$Remitente).Trim().ToLowerInvariant().Split('@')[0]
}

#endregion

#region Documentos --------------------------------------------------------------

<#
.SYNOPSIS
    Extensión en minúsculas con punto. "Factura.PDF" -> ".pdf"; "archivo" -> ".archivo".
.DESCRIPTION
    Expresión equivalente: concat('.', toLower(last(split(item()?['name'], '.'))))
#>
function Get-ExtensionArchivo {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Nombre)
    $partes = ([string]$Nombre).Split('.')
    return '.' + $partes[$partes.Length - 1].ToLowerInvariant()
}

<#
.SYNOPSIS
    Tipo de documento según la extensión: 'factura', 'comprobante', 'xml' o $null (no soportado).
#>
function Get-TipoDocumento {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$NombreArchivo)
    $extension = Get-ExtensionArchivo $NombreArchivo
    foreach ($tipo in Get-TiposDocumento) {
        if ($tipo.Extensiones -contains $extension) { return $tipo.Clave }
    }
    return $null
}

<#
.SYNOPSIS
    Valida los adjuntos de un correo (paso 5 del proceso) y arma los hallazgos para el colaborador.
.PARAMETER Adjuntos
    Nombres de los archivos adjuntos que NO son imágenes incrustadas en el cuerpo del correo.
.PARAMETER Asunto
    Asunto del correo, para avisar si no trae el nombre del colaborador.
.OUTPUTS
    Facturas, Comprobantes, Xmls, NoSoportados, Soportados, Completa y Hallazgos (en el orden
    en que el flujo los agrega a varHallazgos).
#>
function Test-Documentacion {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$Adjuntos = @(),
        [AllowNull()][AllowEmptyString()][string]$Asunto
    )
    $reglas = Get-ReglasDominio
    $porTipo = [ordered]@{}
    foreach ($tipo in Get-TiposDocumento) { $porTipo[$tipo.Clave] = @() }
    $noSoportados = @()
    foreach ($nombre in @($Adjuntos)) {
        $tipoArchivo = Get-TipoDocumento $nombre
        if ($tipoArchivo) { $porTipo[$tipoArchivo] += $nombre } else { $noSoportados += $nombre }
    }

    $hallazgos = @()
    $completa = $true
    foreach ($tipo in Get-TiposDocumento) {
        if ($tipo.Obligatorio -and @($porTipo[$tipo.Clave]).Count -eq 0) {
            $completa = $false
            if ($tipo.Clave -eq 'factura') { $hallazgos += $reglas.mensajes.faltaFactura }
            elseif ($tipo.Clave -eq 'comprobante') { $hallazgos += $reglas.mensajes.faltaComprobante }
            else { $hallazgos += ('Falta: {0}.' -f $tipo.Nombre) }
        }
    }
    if ($noSoportados.Count -gt 0) {
        $hallazgos += $reglas.mensajes.archivosNoSoportados.Replace('{archivos}', ($noSoportados -join ', ')).Replace('{formatos}', ((Get-ExtensionesPermitidas) -join ', '))
    }
    if (-not (Get-NombreEnAsunto $Asunto)) {
        $hallazgos += $reglas.mensajes.asuntoSinNombre
    }

    $soportados = @()
    foreach ($clave in $porTipo.Keys) { $soportados += $porTipo[$clave] }

    [pscustomobject]@{
        Facturas     = @($porTipo['factura'])
        Comprobantes = @($porTipo['comprobante'])
        Xmls         = @($porTipo['xml'])
        NoSoportados = $noSoportados
        Soportados   = $soportados
        Completa     = $completa
        Hallazgos    = $hallazgos
    }
}

#endregion

#region Caso --------------------------------------------------------------------

<#
.SYNOPSIS
    Fecha local a partir de la fecha UTC de recepción del correo.
.PARAMETER ZonaHoraria
    Id. de zona horaria de Windows, el mismo que usa convertFromUtc en Power Automate
    (México: 'Central Standard Time (Mexico)'; Colombia: 'SA Pacific Standard Time').
#>
function ConvertTo-FechaLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$FechaUtc,
        [Parameter(Mandatory)][string]$ZonaHoraria
    )
    $utc = [datetime]::SpecifyKind($FechaUtc, [System.DateTimeKind]::Utc)
    return [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, [System.TimeZoneInfo]::FindSystemTimeZoneById($ZonaHoraria))
}

<#
.SYNOPSIS
    Identificador del caso: PREFIJO-AAAAMM-NNNNN (ej. LEG-202609-00123).
.DESCRIPTION
    AAAAMM es el mes LOCAL de recepción; NNNNN es el ID interno del elemento de SharePoint.
    Expresión equivalente (acción Id_legalizacion):
      concat(outputs('Configuracion')?['prefijoId'], '-', outputs('Anio'), outputs('Mes'), '-',
             formatNumber(outputs('Crear_registro')?['body/ID'], '00000'))
#>
function New-IdLegalizacion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetime]$FechaRecepcionLocal,
        [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$Consecutivo
    )
    $reglas = Get-ReglasDominio
    $formato = '{0}-{1:yyyyMM}-{2:D' + $reglas.caso.digitosConsecutivo + '}'
    return ($formato -f $reglas.caso.prefijoId, $FechaRecepcionLocal, $Consecutivo)
}

<#
.SYNOPSIS
    Año y mes (AAAA, MM) codificados en un Id. de legalización. Los usa el flujo de cierre.
#>
function Get-PeriodoDesdeId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$IdLegalizacion)
    $prefijo = (Get-ReglasDominio).caso.prefijoId
    $inicio = $prefijo.Length + 1
    [pscustomobject]@{
        Anio = $IdLegalizacion.Substring($inicio, 4)
        Mes  = $IdLegalizacion.Substring($inicio + 4, 2)
    }
}

<#
.SYNOPSIS
    Nombre de carpeta seguro para SharePoint a partir del nombre del colaborador.
.DESCRIPTION
    Quita los caracteres no permitidos del dominio, recorta espacios, limita el largo y,
    si queda vacío, usa el nombre por defecto. Expresión equivalente en expresiones.md (E-07).
#>
function ConvertTo-NombreCarpeta {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Texto)
    $carpetas = (Get-ReglasDominio).carpetas
    $resultado = [string]$Texto
    foreach ($caracter in $carpetas.caracteresNoPermitidos) { $resultado = $resultado.Replace([string]$caracter, '') }
    $resultado = $resultado.Trim()
    if ($resultado.Length -gt [int]$carpetas.largoMaximoNombre) { $resultado = $resultado.Substring(0, [int]$carpetas.largoMaximoNombre) }
    $resultado = $resultado.Trim()
    if (-not $resultado) { $resultado = $carpetas.nombreSinDato }
    return $resultado
}

<#
.SYNOPSIS
    Carpeta del caso dentro de la biblioteca: Raiz/AAAA/MM/Colaborador (ej. Pendientes/2026/09/Ana Pérez).
#>
function Get-RutaCarpetaCaso {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Estado,
        [Parameter(Mandatory)][datetime]$FechaRecepcionLocal,
        [AllowNull()][AllowEmptyString()][string]$Colaborador
    )
    $raices = (Get-ReglasDominio).carpetas.raizPorEstado
    if (-not $raices.PSObject.Properties[$Estado]) { throw "Estado desconocido: '$Estado'." }
    return '{0}/{1:yyyy}/{1:MM}/{2}' -f $raices.$Estado, $FechaRecepcionLocal, (ConvertTo-NombreCarpeta $Colaborador)
}

<#
.SYNOPSIS
    ¿Está permitido pasar de un estado a otro? Aprobada y Rechazada son finales.
#>
function Test-TransicionEstado {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Desde,
        [Parameter(Mandatory)][string]$Hacia
    )
    $caso = (Get-ReglasDominio).caso
    if (@($caso.estados) -notcontains $Desde) { throw "Estado desconocido: '$Desde'." }
    if (@($caso.estados) -notcontains $Hacia) { throw "Estado desconocido: '$Hacia'." }
    if ($Desde -eq $Hacia) { return $true }
    return @($caso.transiciones.$Desde) -contains $Hacia
}

#endregion

#region Duplicados --------------------------------------------------------------

<#
.SYNOPSIS
    Casos previos que podrían ser el mismo envío (heurística, no revisa el contenido).
.DESCRIPTION
    Un caso previo es posible duplicado si:
      - es del mismo remitente (sin distinguir mayúsculas),
      - llegó dentro de los últimos duplicados.diasVentana días,
      - su estado no está en duplicados.estadosExcluidos (un reenvío tras un rechazo es normal),
      - y comparte al menos un nombre de archivo (sin distinguir mayúsculas) de un tipo que
        compara duplicados (PDF y XML; el Excel quincenal suele llamarse igual siempre).
.PARAMETER CasosPrevios
    Objetos con IdLegalizacion, Correo, FechaRecepcionUtc, Estado y NombreArchivo
    (nombres separados por salto de línea, como se guardan en SharePoint).
.OUTPUTS
    Ids de legalización de los casos que coinciden, en el orden recibido.
#>
function Find-PosiblesDuplicados {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][string[]]$NuevosArchivos = @(),
        [Parameter(Mandatory)][string]$Remitente,
        [Parameter(Mandatory)][datetime]$FechaRecepcionUtc,
        [AllowEmptyCollection()][object[]]$CasosPrevios = @()
    )
    $reglas = Get-ReglasDominio
    $tiposComparables = @(Get-TiposDocumento | Where-Object { $_.ComparaDuplicados } | ForEach-Object { $_.Clave })
    $nombresComparables = @($NuevosArchivos | Where-Object { $tiposComparables -contains (Get-TipoDocumento $_) } | ForEach-Object { $_.ToLowerInvariant() })
    if ($nombresComparables.Count -eq 0) { return @() }

    $desde = $FechaRecepcionUtc.AddDays(-[int]$reglas.duplicados.diasVentana)
    $excluidos = @($reglas.duplicados.estadosExcluidos)
    $resultado = @()
    foreach ($caso in @($CasosPrevios)) {
        if ([string]$caso.Correo -ne $Remitente) { continue }            # -ne no distingue mayúsculas
        if ([datetime]$caso.FechaRecepcionUtc -lt $desde) { continue }
        if ($excluidos -contains [string]$caso.Estado) { continue }
        $nombresCaso = @(([string]$caso.NombreArchivo).ToLowerInvariant().Replace("`r", '').Split("`n"))
        $comunes = @($nombresCaso | Where-Object { $nombresComparables -contains $_ })
        if ($comunes.Count -gt 0) { $resultado += [string]$caso.IdLegalizacion }
    }
    return $resultado
}

<#
.SYNOPSIS
    Texto del aviso de posible duplicado que queda en Observaciones y en el correo a Contabilidad.
#>
function Get-MensajePosibleDuplicado {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Casos)
    $reglas = Get-ReglasDominio
    return $reglas.mensajes.posibleDuplicado.Replace('{casos}', ($Casos -join ', ')).Replace('{dias}', [string]$reglas.duplicados.diasVentana)
}

#endregion

#region Observaciones -----------------------------------------------------------

<#
.SYNOPSIS
    Líneas de Observaciones que se pueden mostrar al colaborador: quita las vacías y las notas
    internas (las que empiezan con observaciones.prefijoNotaInterna, p. ej. "[Interno]").
.DESCRIPTION
    Las usa LEG-02 para el correo de cierre. Expresión equivalente (acción Observaciones_para_colaborador,
    Filtrar matriz desde split(coalesce(triggerBody()?['Observaciones'], ''), decodeUriComponent('%0A'))):
      @and(not(empty(trim(item()))), not(startsWith(trim(item()), outputs('Configuracion')?['prefijoNotaInterna'])))
#>
function Get-ObservacionesParaColaborador {
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Observaciones)
    $prefijo = (Get-ReglasDominio).observaciones.prefijoNotaInterna
    $lineas = @()
    foreach ($linea in ([string]$Observaciones).Split("`n")) {
        $limpia = $linea.Trim()
        if ($limpia -and -not $limpia.StartsWith($prefijo, [System.StringComparison]::Ordinal)) { $lineas += $limpia }
    }
    return $lineas
}

#endregion

Export-ModuleMember -Function @(
    'Import-ReglasDominio', 'Get-ReglasDominio', 'Get-TiposDocumento', 'Get-ExtensionesPermitidas',
    'ConvertTo-AsuntoNormalizado', 'Test-AsuntoLegalizacion', 'Get-NombreEnAsunto', 'Get-NombreColaborador',
    'Get-ExtensionArchivo', 'Get-TipoDocumento', 'Test-Documentacion',
    'ConvertTo-FechaLocal', 'New-IdLegalizacion', 'Get-PeriodoDesdeId', 'ConvertTo-NombreCarpeta', 'Get-RutaCarpetaCaso', 'Test-TransicionEstado',
    'Find-PosiblesDuplicados', 'Get-MensajePosibleDuplicado', 'Get-ObservacionesParaColaborador'
)
