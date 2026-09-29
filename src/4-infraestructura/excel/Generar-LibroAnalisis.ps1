<#
.SYNOPSIS
    Genera el libro de análisis de Excel (Power Query) conectado al registro de SharePoint.

.DESCRIPTION
    Capa de INFRAESTRUCTURA (Excel). Arma el libro a partir de:
      - las consultas M de src/3-adaptadores/excel/powerquery (adaptador),
      - las reglas de src/1-dominio/dominio.json (estados, tipos de documento, prefijo),
      - el contrato de SharePoint (Esquema.Legalizaciones.psd1) y la configuración del entorno.

    Usa Excel de escritorio por automatización COM porque es la única forma de crear
    consultas de Power Query dentro de un archivo (openpyxl y similares no pueden).
    El libro generado se sube luego a SharePoint y se abre con Excel de escritorio o web.

    Proceso: crea hojas y consultas, las ejecuta con datos de DEMOSTRACIÓN para dar formato,
    validaciones y fórmulas a las tablas, y al final:
      - Modo SharePoint: vacía las tablas y deja pModo = "SharePoint" (se llenan al actualizar).
      - Modo Demostracion: deja los datos de ejemplo para conocer el libro sin SharePoint.

    Requiere Microsoft Excel de escritorio (Microsoft 365 Apps). Funciona en Windows PowerShell 5.1 y PowerShell 7.

.PARAMETER Modo
    SharePoint (predeterminado): exige Tenant.Nombre completo en la configuración.
    Demostracion: no necesita configuración de tenant.

.PARAMETER RutaSalida
    Archivo a generar. Por defecto salida/Analisis_Legalizaciones_TC.xlsx (o _DEMO.xlsx).

.EXAMPLE
    ./src/4-infraestructura/excel/Generar-LibroAnalisis.ps1

.EXAMPLE
    ./src/4-infraestructura/excel/Generar-LibroAnalisis.ps1 -Modo Demostracion
#>
[CmdletBinding()]
param(
    [ValidateSet('SharePoint', 'Demostracion')][string]$Modo = 'SharePoint',
    [string]$RutaConfig,
    [string]$RutaSalida
)

$ErrorActionPreference = 'Stop'
$aqui = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$raiz = (Resolve-Path (Join-Path $aqui '..\..\..')).Path
if (-not $RutaConfig) { $RutaConfig = Join-Path $raiz 'config\Config.Legalizaciones.psd1' }
. (Join-Path $raiz 'src\4-infraestructura\sharepoint\Comun.Legalizaciones.ps1')

#region Entradas -----------------------------------------------------------------

function Read-Json([string]$Ruta) {
    return [System.IO.File]::ReadAllText($Ruta, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
}
function ConvertTo-ListaM([object[]]$Valores) {
    return '{' + ((@($Valores) | ForEach-Object { '"' + ([string]$_).Replace('"', '""') + '"' }) -join ', ') + '}'
}

$cfg = Import-PowerShellDataFile -LiteralPath $RutaConfig
$esq = Import-PowerShellDataFile -LiteralPath (Join-Path $raiz 'src\3-adaptadores\sharepoint\Esquema.Legalizaciones.psd1')
$dominio = Read-Json (Join-Path $raiz 'src\1-dominio\dominio.json')
$plantilla = Read-Json (Join-Path $raiz 'src\3-adaptadores\excel\plantilla-comprobante.json')
$carpetaConsultas = Join-Path $raiz 'src\3-adaptadores\excel\powerquery'
$region = Get-RegionEntorno -Config $cfg

$tenantPendiente = @(Find-Marcadores -Valor $cfg.Tenant.Nombre -Ruta 'Config.Tenant.Nombre').Count -gt 0
if ($Modo -eq 'SharePoint' -and $tenantPendiente) {
    throw "Para el modo SharePoint complete Tenant.Nombre en '$RutaConfig'. Para conocer el libro sin SharePoint use: -Modo Demostracion"
}
$tenant = if ($tenantPendiente) { 'ejemplo' } else { ([string]$cfg.Tenant.Nombre).ToLowerInvariant() }
$sitioUrl = 'https://{0}.sharepoint.com{1}' -f $tenant, $cfg.Sitio.RutaRelativa

if (-not $RutaSalida) {
    $nombre = if ($Modo -eq 'Demostracion') { 'Analisis_Legalizaciones_TC_DEMO.xlsx' } else { 'Analisis_Legalizaciones_TC.xlsx' }
    $RutaSalida = Join-Path $raiz "salida\$nombre"
}
$RutaSalida = [System.IO.Path]::GetFullPath($RutaSalida)
if ($RutaSalida.Length -gt 200) { throw "La ruta de salida es demasiado larga para Excel ($($RutaSalida.Length) caracteres; máximo 200): $RutaSalida" }
New-Item -ItemType Directory -Path (Split-Path $RutaSalida) -Force | Out-Null

$tiposDocumento = '[' + ((@($dominio.documentos.PSObject.Properties) | ForEach-Object { '{0} = {1}' -f $_.Name, (ConvertTo-ListaM $_.Value.extensiones) }) -join ', ') + ']'
$tokens = [ordered]@{
    SITIO_URL             = $sitioUrl
    LISTA_REGISTRO        = $esq.Lista.Titulo
    RUTA_BIBLIOTECA       = '/{0}/' -f $esq.Biblioteca.Url
    DESFASE_UTC           = [string]$region.DesfaseHorasUTC
    DIAS_ALERTA           = [string]$dominio.analisis.diasAlertaCasoAbierto
    MODO                  = 'Demostracion'
    PREFIJO_ID            = $dominio.caso.prefijoId
    ESTADOS               = ConvertTo-ListaM $dominio.caso.estados
    ESTADOS_ABIERTOS      = ConvertTo-ListaM $dominio.caso.estadosAbiertos
    ESTADOS_FINALES       = ConvertTo-ListaM $dominio.caso.estadosFinales
    ESTADO_APROBADO       = $dominio.caso.estadoAprobado
    TIPOS_DOCUMENTO       = $tiposDocumento
    TEXTO_AVISO_DUPLICADO = $dominio.mensajes.posibleDuplicado.Split('{')[0].Trim()
    URL_LISTA             = $esq.Lista.Url
    COLUMNAS_PLANTILLA    = ConvertTo-ListaM @($plantilla.columnas | ForEach-Object { $_.nombre })
    TABLA_PLANTILLA       = $plantilla.tabla
}

# Consultas en orden de creación, con su descripción.
$consultas = [ordered]@{
    pSitioUrl          = 'Parámetro: URL del sitio de SharePoint.'
    pListaRegistro     = 'Parámetro: nombre de la lista del registro.'
    pRutaBiblioteca    = 'Parámetro: ruta de la biblioteca de soportes.'
    pDesfaseHorasUTC   = 'Parámetro: horas de diferencia entre la hora local y UTC.'
    pDiasAlertaAbierto = 'Parámetro: días para alertar un caso abierto.'
    pModo              = 'Parámetro: SharePoint (datos reales), Demostracion o Vacio.'
    Dominio            = 'Reglas del dominio (generadas desde dominio.json).'
    fnUtcALocal        = 'Función: fecha UTC a hora local.'
    fnLeerComprobante  = 'Función: lee la tabla de gastos de un comprobante quincenal.'
    Registro           = 'Registro de legalizaciones (lista de SharePoint).'
    SoportesBase       = 'Archivos de la biblioteca con contenido (no se carga).'
    Soportes           = 'Archivos de la biblioteca por caso.'
    Gastos             = 'Líneas de gasto de los comprobantes quincenales.'
    Gastos_Errores     = 'Comprobantes que no se pudieron leer.'
    Control_Operam     = 'Casos ya cargados en Operam (tabla manual).'
    Consolidado_Operam = 'Gastos aprobados pendientes de cargar en Operam.'
    Alertas            = 'Controles de calidad.'
    Resumen_Mensual    = 'Casos por mes y estado.'
}

# Consultas que se cargan a una tabla, en orden de actualización.
$cargas = @(
    @{ Consulta = 'Registro';           Hoja = 'Registro';           Tabla = 'tblRegistro';           Celda = 'A1' }
    @{ Consulta = 'Soportes';           Hoja = 'Soportes';           Tabla = 'tblSoportes';           Celda = 'A1' }
    @{ Consulta = 'Gastos';             Hoja = 'Gastos';             Tabla = 'tblGastosConsolidados'; Celda = 'A1' }
    @{ Consulta = 'Gastos_Errores';     Hoja = 'Errores_Lectura';    Tabla = 'tblErroresLectura';     Celda = 'A1' }
    @{ Consulta = 'Consolidado_Operam'; Hoja = 'Consolidado_Operam'; Tabla = 'tblConsolidadoOperam';  Celda = 'A1' }
    @{ Consulta = 'Alertas';            Hoja = 'Alertas';            Tabla = 'tblAlertas';            Celda = 'A1' }
    @{ Consulta = 'Resumen_Mensual';    Hoja = 'Resumen';            Tabla = 'tblResumenMensual';     Celda = 'A8' }
)
$hojas = @('Inicio', 'Resumen', 'Registro', 'Alertas', 'Consolidado_Operam', 'Control_Operam', 'Gastos', 'Soportes', 'Errores_Lectura', 'Listas')

function Get-FormulaConsulta([string]$Nombre) {
    $ruta = Join-Path $carpetaConsultas "$Nombre.pq"
    $texto = [System.IO.File]::ReadAllText($ruta, [System.Text.Encoding]::UTF8)
    foreach ($clave in $tokens.Keys) { $texto = $texto.Replace('{{' + $clave + '}}', [string]$tokens[$clave]) }
    if ($texto -match '\{\{[A-Z_]+\}\}') { throw "La consulta $Nombre tiene un marcador sin reemplazar: $($Matches[0])" }
    return $texto.Trim()
}

#endregion

#region Excel ---------------------------------------------------------------------

$COLOR = @{ Aprobada = 13561798; Rechazada = 13551615; Pendiente = 10284031; 'En revisión' = 16247773; Encabezado = 6299648 }

function Set-FormatoColumna($Tabla, [string]$Columna, [string]$Formato) {
    $rango = $Tabla.ListColumns.Item($Columna).DataBodyRange
    if ($rango) { $rango.NumberFormat = $Formato }
}

function Set-AnchoColumnas($Hoja) {
    $usado = $Hoja.UsedRange
    [void]$usado.Columns.AutoFit()
    foreach ($columna in $usado.Columns) { if ($columna.ColumnWidth -gt 55) { $columna.ColumnWidth = 55 } }
}

function Set-EncabezadoFijo($Libro, $Hoja, [int]$Fila) {
    try {
        $Hoja.Activate()
        $ventana = $Libro.Windows.Item(1)
        $ventana.FreezePanes = $false
        $ventana.SplitColumn = 0
        $ventana.SplitRow = $Fila
        $ventana.FreezePanes = $true
    }
    catch { Write-Verbose "No se pudo inmovilizar la fila de encabezado de $($Hoja.Name): $($_.Exception.Message)" }
}

function Update-TablaConsulta($Tabla, [string]$Consulta) {
    try { [void]$Tabla.QueryTable.Refresh($false) }
    catch { throw "Falló la actualización de la consulta '$Consulta' (tabla $($Tabla.Name)): $($_.Exception.Message)" }
}

#endregion

Write-Host "Generando libro de análisis (modo $Modo)" -ForegroundColor Cyan
Write-Host "  Sitio: $sitioUrl"
Write-Host "  País: $($region.Pais) (UTC$($region.DesfaseHorasUTC))"

$excel = $null
$libro = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $libro = $excel.Workbooks.Add()
    # Hojas en orden (un libro nuevo puede traer varias hojas según la configuración de Excel)
    while ($libro.Worksheets.Count -gt 1) { $libro.Worksheets.Item(2).Delete() }
    $libro.Worksheets.Item(1).Name = $hojas[0]
    for ($i = 1; $i -lt $hojas.Count; $i++) {
        $nueva = $libro.Worksheets.Add([Type]::Missing, $libro.Worksheets.Item($libro.Worksheets.Count))
        $nueva.Name = $hojas[$i]
    }

    # Listas del dominio (lista desplegable de Estado)
    $listas = $libro.Worksheets.Item('Listas')
    $listas.Range('A1').Value2 = 'Estados'
    $fila = 2
    foreach ($estado in $dominio.caso.estados) { $listas.Cells.Item($fila, 1).Value2 = $estado; $fila++ }
    [void]$libro.Names.Add('lstEstados', ('=Listas!$A$2:$A${0}' -f ($fila - 1)))
    # Las validaciones de datos por COM interpretan las fórmulas en el idioma de Excel instalado:
    # se usan nombres definidos (sin funciones ni separadores) para que funcione en cualquier idioma.
    $listas.Range('B1').Value2 = 'Fecha mínima de carga'
    $listas.Range('B2').Value2 = [double]([datetime]'2020-01-01').ToOADate()
    [void]$libro.Names.Add('fechaMinimaCarga', '=Listas!$B$2')
    $listas.Visible = 0   # xlSheetHidden

    # Tabla manual de control de cargas a Operam
    $control = $libro.Worksheets.Item('Control_Operam')
    $encabezados = @('ID Legalización', 'Fecha Carga', 'Cargado Por', 'Referencia Operam', 'Observaciones')
    for ($c = 0; $c -lt $encabezados.Count; $c++) { $control.Cells.Item(1, $c + 1).Value2 = $encabezados[$c] }
    $tablaControl = $control.ListObjects.Add(1, $control.Range('A1:E2'), [Type]::Missing, 1)
    $tablaControl.Name = 'tblControlOperam'
    $tablaControl.TableStyle = 'TableStyleMedium4'
    if ($Modo -eq 'Demostracion') {
        $control.Range('A2').Value2 = 'LEG-202608-00095'
        $control.Range('B2').Formula = '=TODAY()-25'
        $control.Range('C2').Value2 = 'Contabilidad (DEMO)'
        $control.Range('D2').Value2 = 'OP-2026-0815'
        $control.Range('E2').Value2 = 'Ejemplo: caso ya cargado; no aparece en Consolidado_Operam.'
    }
    $validacionFecha = $tablaControl.ListColumns.Item('Fecha Carga').DataBodyRange.Validation
    $validacionFecha.Delete()
    $validacionFecha.Add(4, 1, 7, '=fechaMinimaCarga')   # xlValidateDate, xlGreaterEqual
    $validacionFecha.ErrorMessage = 'Escriba la fecha en que se cargó el caso en Operam.'
    $tablaControl.ListColumns.Item('Fecha Carga').DataBodyRange.NumberFormat = 'dd/mm/yyyy'

    # Consultas de Power Query
    foreach ($nombre in $consultas.Keys) {
        [void]$libro.Queries.Add($nombre, (Get-FormulaConsulta $nombre), $consultas[$nombre])
    }
    Write-Host "  Consultas creadas: $($libro.Queries.Count)"

    # Tablas conectadas a las consultas
    $tablas = @{}
    foreach ($carga in $cargas) {
        $hoja = $libro.Worksheets.Item($carga.Hoja)
        $conexion = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={0};Extended Properties=""' -f $carga.Consulta
        $tabla = $hoja.ListObjects.Add(0, $conexion, [Type]::Missing, [Type]::Missing, $hoja.Range($carga.Celda))
        $qt = $tabla.QueryTable
        $qt.CommandType = 2                       # xlCmdSql
        $qt.CommandText = 'SELECT * FROM [{0}]' -f $carga.Consulta
        $qt.RowNumbers = $false
        $qt.FillAdjacentFormulas = $false
        $qt.PreserveFormatting = $true
        $qt.RefreshOnFileOpen = $false
        $qt.BackgroundQuery = $true
        $qt.RefreshStyle = 1                      # xlInsertDeleteCells
        $qt.SavePassword = $false
        $qt.SaveData = $true
        $qt.AdjustColumnWidth = $false
        $qt.PreserveColumnInfo = $true
        $tabla.Name = $carga.Tabla
        $tabla.TableStyle = 'TableStyleMedium2'
        Update-TablaConsulta $tabla $carga.Consulta
        $tablas[$carga.Tabla] = $tabla
        Write-Host ("  {0,-22} {1,3} filas de demostración" -f $carga.Tabla, $tabla.ListRows.Count)
    }

    # Control_Operam: el ID debe existir en el registro (requiere que tblRegistro ya exista)
    $validacionId = $tablaControl.ListColumns.Item('ID Legalización').DataBodyRange.Validation
    $validacionId.Delete()
    [void]$libro.Names.Add('lstIdsRegistro', '=tblRegistro[ID Legalización]')
    $validacionId.Add(3, 1, 1, '=lstIdsRegistro')
    $validacionId.ErrorTitle = 'ID no encontrado'
    $validacionId.ErrorMessage = 'Escriba un ID Legalización que exista en la hoja Registro.'

    # Formatos, validaciones y formato condicional (se conservan cuando la consulta se actualiza)
    $registro = $tablas['tblRegistro']
    Set-FormatoColumna $registro 'Fecha Recepción' 'dd/mm/yyyy hh:mm'
    Set-FormatoColumna $registro 'Fecha Cierre' 'dd/mm/yyyy hh:mm'
    $estadoRango = $registro.ListColumns.Item('Estado').DataBodyRange
    $estadoRango.Validation.Delete()
    $estadoRango.Validation.Add(3, 1, 1, '=lstEstados')
    $estadoRango.Validation.InputTitle = 'Estado'
    $estadoRango.Validation.InputMessage = 'El estado oficial se cambia en SharePoint; esta tabla se sobrescribe al actualizar.'
    $estadoRango.Validation.ErrorMessage = 'Use uno de los estados de la lista.'
    foreach ($estado in $dominio.caso.estados) {
        if ($COLOR.ContainsKey($estado)) {
            $condicion = $estadoRango.FormatConditions.Add(1, 3, ('="{0}"' -f $estado))
            $condicion.Interior.Color = $COLOR[$estado]
        }
    }
    $consolidado = $tablas['tblConsolidadoOperam']
    Set-FormatoColumna $consolidado 'Fecha Cierre' 'dd/mm/yyyy hh:mm'
    Set-FormatoColumna $consolidado 'Fecha' 'dd/mm/yyyy'
    $gastos = $tablas['tblGastosConsolidados']
    Set-FormatoColumna $gastos 'Fecha' 'dd/mm/yyyy'
    foreach ($columna in 'Subtotal', 'IVA', 'Total') {
        Set-FormatoColumna $consolidado $columna '#,##0.00'
        Set-FormatoColumna $gastos $columna '#,##0.00'
    }
    Set-FormatoColumna $gastos 'Diferencia' '#,##0.00'
    Set-FormatoColumna $tablas['tblSoportes'] 'Fecha Creación' 'dd/mm/yyyy hh:mm'

    # Hoja Resumen: indicadores sobre las tablas
    $resumen = $libro.Worksheets.Item('Resumen')
    $resumen.Range('A1').Value2 = 'Resumen de legalizaciones'
    $resumen.Range('A1').Font.Size = 16
    $resumen.Range('A1').Font.Bold = $true
    $abiertos = (@($dominio.caso.estadosAbiertos) | ForEach-Object { '"' + $_ + '"' }) -join ','
    $indicadores = @(
        @('Casos en el registro', '=COUNTA(tblRegistro[ID Legalización])'),
        @('Casos abiertos', ('=SUM(COUNTIFS(tblRegistro[Estado],{{{0}}}))' -f $abiertos)),
        @('Líneas de gasto por cargar a Operam', '=COUNTA(tblConsolidadoOperam[ID Legalización])'),
        @('Alertas', '=COUNTA(tblAlertas[Alerta])')
    )
    $fila = 3
    foreach ($indicador in $indicadores) {
        $resumen.Cells.Item($fila, 1).Value2 = $indicador[0]
        $resumen.Cells.Item($fila, 2).Formula = $indicador[1]
        $resumen.Cells.Item($fila, 2).Font.Bold = $true
        $fila++
    }
    $resumen.Range('A7').Value2 = 'Casos por mes de recepción y estado'
    $resumen.Range('A7').Font.Bold = $true

    # Hoja Inicio
    $inicio = $libro.Worksheets.Item('Inicio')
    $lineas = @(
        'Análisis de legalizaciones de tarjetas corporativas',
        ('Emobility · Fase 1 · Generado el {0:dd/MM/yyyy HH:mm} · Modo: {1}' -f (Get-Date), $Modo),
        '',
        'Cómo usarlo',
        '1. Datos > Actualizar todo. La primera vez, inicie sesión con su cuenta de Microsoft 365 (Cuenta organizacional).',
        '2. Revise las hojas Alertas y Resumen.',
        '3. Consolidado_Operam muestra los gastos de los casos Aprobados que todavía no se cargaron en Operam.',
        '4. Después de cargarlos en Operam, anote cada ID en la hoja Control_Operam y vuelva a actualizar: desaparecen del consolidado.',
        '5. El estado de cada caso se cambia en SharePoint (lista Registro_Legalizaciones_TC), no en este libro.',
        '',
        'Hojas',
        'Resumen: indicadores y casos por mes.   Registro: todos los casos.   Alertas: controles de calidad.',
        'Consolidado_Operam: gastos por cargar.   Control_Operam: registro manual de lo cargado.',
        'Gastos: líneas de todos los comprobantes.   Soportes: archivos por caso.   Errores_Lectura: comprobantes que no usan la plantilla.',
        '',
        'Parámetros (Datos > Consultas y conexiones)',
        ('pSitioUrl = {0}' -f $sitioUrl),
        ('pListaRegistro = {0}   ·   pRutaBiblioteca = {1}' -f $tokens.LISTA_REGISTRO, $tokens.RUTA_BIBLIOTECA),
        ('pDesfaseHorasUTC = {0} ({1})   ·   pDiasAlertaAbierto = {2}' -f $tokens.DESFASE_UTC, $region.Pais, $tokens.DIAS_ALERTA),
        ('pModo = {0}   (SharePoint: datos reales · Demostracion: datos de ejemplo · Vacio: sin datos)' -f $(if ($Modo -eq 'Demostracion') { 'Demostracion' } else { 'SharePoint' }))
    )
    for ($i = 0; $i -lt $lineas.Count; $i++) { $inicio.Cells.Item($i + 1, 1).Value2 = $lineas[$i] }
    $inicio.Range('A1').Font.Size = 16
    $inicio.Range('A1').Font.Bold = $true
    foreach ($celda in 'A4', 'A11', 'A16') { $inicio.Range($celda).Font.Bold = $true }
    $inicio.Columns.Item(1).ColumnWidth = 120

    # Anchos y encabezados fijos
    foreach ($nombreHoja in 'Registro', 'Alertas', 'Consolidado_Operam', 'Control_Operam', 'Gastos', 'Soportes', 'Errores_Lectura') {
        $hoja = $libro.Worksheets.Item($nombreHoja)
        Set-AnchoColumnas $hoja
        Set-EncabezadoFijo $libro $hoja 1
    }
    $resumen.Columns.Item(1).ColumnWidth = 40

    # Modo final
    if ($Modo -eq 'SharePoint') {
        $libro.Queries.Item('pModo').Formula = (Get-FormulaConsulta 'pModo').Replace('"Demostracion" meta', '"Vacio" meta')
        foreach ($carga in $cargas) { Update-TablaConsulta $tablas[$carga.Tabla] $carga.Consulta }
        $libro.Queries.Item('pModo').Formula = (Get-FormulaConsulta 'pModo').Replace('"Demostracion" meta', '"SharePoint" meta')
        Write-Host '  Tablas vaciadas; pModo = "SharePoint" (se llenan con Datos > Actualizar todo).'
    }

    $inicio.Activate()
    if (Test-Path -LiteralPath $RutaSalida) { Remove-Item -LiteralPath $RutaSalida -Force }
    $libro.SaveAs($RutaSalida, 51)   # xlOpenXMLWorkbook
    Write-Host "Libro generado: $RutaSalida" -ForegroundColor Green
}
finally {
    if ($libro) { try { $libro.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
