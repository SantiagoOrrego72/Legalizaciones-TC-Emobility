<#
    Utilidades para probar Power Query con Excel de escritorio (COM) sin SharePoint.
    Las usan pruebas/unitarias/Excel.Tests.ps1.
#>

function Test-ExcelDisponible {
    try {
        $clsid = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CLSID' -ErrorAction Stop).'(default)'
        return [bool]$clsid
    }
    catch { return $false }
}

# Evalúa una consulta M en un libro temporal y devuelve sus filas (como matriz de valores) o el error.
function Invoke-ConsultaM {
    param(
        [Parameter(Mandatory)][hashtable]$Consultas,   # nombre -> fórmula M (todas se agregan al libro)
        [Parameter(Mandatory)][string]$Cargar           # consulta que se carga a la hoja
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $libro = $null
    try {
        $libro = $excel.Workbooks.Add()
        foreach ($nombre in $Consultas.Keys) { [void]$libro.Queries.Add($nombre, $Consultas[$nombre]) }
        $hoja = $libro.Worksheets.Item(1)
        $conexion = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={0};Extended Properties=""' -f $Cargar
        $tabla = $hoja.ListObjects.Add(0, $conexion, [Type]::Missing, [Type]::Missing, $hoja.Range('A1'))
        $tabla.QueryTable.CommandType = 2
        $tabla.QueryTable.CommandText = 'SELECT * FROM [{0}]' -f $Cargar
        try {
            [void]$tabla.QueryTable.Refresh($false)
        }
        catch {
            return [pscustomobject]@{ Error = $_.Exception.Message; Filas = 0; Encabezados = @(); Valores = @() }
        }
        $encabezados = @($tabla.HeaderRowRange.Value2)
        $valores = @()
        if ($tabla.ListRows.Count -gt 0) { $valores = $tabla.DataBodyRange.Value2 }
        return [pscustomobject]@{ Error = $null; Filas = $tabla.ListRows.Count; Encabezados = $encabezados; Valores = $valores }
    }
    finally {
        if ($libro) { try { $libro.Close($false) } catch { } }
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Lista de nombres de consulta de un libro y la fórmula de pModo (abre el libro sin actualizarlo).
function Get-ConsultasLibro {
    param([Parameter(Mandatory)][string]$Ruta)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $libro = $null
    try {
        $libro = $excel.Workbooks.Open($Ruta, 0, $true)
        $consultas = @()
        foreach ($consulta in $libro.Queries) { $consultas += $consulta.Name }
        $tablas = @{}
        foreach ($hoja in $libro.Worksheets) {
            foreach ($tabla in $hoja.ListObjects) { $tablas[$tabla.Name] = [pscustomobject]@{ Hoja = $hoja.Name; Filas = $tabla.ListRows.Count; Encabezados = @($tabla.HeaderRowRange.Value2) } }
        }
        return [pscustomobject]@{
            Consultas = $consultas
            Modo      = $libro.Queries.Item('pModo').Formula
            Tablas    = $tablas
            Nombres   = @(foreach ($n in $libro.Names) { $n.Name })
        }
    }
    finally {
        if ($libro) { try { $libro.Close($false) } catch { } }
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
