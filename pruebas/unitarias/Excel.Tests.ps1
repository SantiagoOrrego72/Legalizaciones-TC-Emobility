<#
    Pruebas del adaptador de Excel (Pester 5). Ejecutan Power Query de verdad con Excel de escritorio (COM),
    sin SharePoint: el modo "Demostracion" de las consultas reemplaza los orígenes de datos.
    Se omiten si Excel no está instalado.
#>

BeforeDiscovery {
    . (Join-Path $PSScriptRoot 'ExcelPrueba.Comun.ps1')
    $script:SinExcel = -not (Test-ExcelDisponible)
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'ExcelPrueba.Comun.ps1')
    $script:Raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    $script:Plantilla = Get-Content (Join-Path $Raiz 'src\3-adaptadores\excel\plantilla-comprobante.json') -Raw -Encoding utf8 | ConvertFrom-Json
    $script:Columnas = @($Plantilla.columnas | ForEach-Object { $_.nombre })
    $listaM = '{' + (($Columnas | ForEach-Object { '"' + $_ + '"' }) -join ', ') + '}'
    $script:FnLeer = (Get-Content (Join-Path $Raiz 'src\3-adaptadores\excel\powerquery\fnLeerComprobante.pq') -Raw -Encoding utf8).
        Replace('{{COLUMNAS_PLANTILLA}}', $listaM).Replace('{{TABLA_PLANTILLA}}', $Plantilla.tabla)

    # Datos de prueba recién generados (Python + openpyxl) en una carpeta temporal
    $script:Datos = Join-Path $TestDrive 'datos'
    & python (Join-Path $Raiz 'pruebas\aceptacion\generar_datos_prueba.py') --destino $Datos | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'No se pudieron generar los datos de prueba con Python.' }

    $script:Generador = Join-Path $Raiz 'src\4-infraestructura\excel\Generar-LibroAnalisis.ps1'
    $script:CarpetaCorta = Join-Path ([System.IO.Path]::GetTempPath()) ('leg-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $CarpetaCorta -Force | Out-Null
}

AfterAll {
    if ($CarpetaCorta -and (Test-Path $CarpetaCorta)) { Remove-Item $CarpetaCorta -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Contrato entre la plantilla del comprobante y Power Query (fnLeerComprobante)' -Skip:$SinExcel {
    It 'lee las 3 líneas del comprobante hecho con la plantilla, con las columnas del contrato' {
        $ruta = (Join-Path $Datos 'Gastos_Quincena.xlsx').Replace('"', '""')
        $r = Invoke-ConsultaM -Consultas @{ fnLeerComprobante = $FnLeer; Leer = "fnLeerComprobante(File.Contents(""$ruta""))" } -Cargar 'Leer'
        $r.Error | Should -BeNullOrEmpty
        $r.Filas | Should -Be 3
        $r.Encabezados | Should -Be $Columnas
    }
    It 'falla con un Excel de formato libre (irá a Errores_Lectura)' {
        $ruta = (Join-Path $Datos 'Gastos_Formato_Libre.xlsx').Replace('"', '""')
        $r = Invoke-ConsultaM -Consultas @{ fnLeerComprobante = $FnLeer; Leer = "fnLeerComprobante(File.Contents(""$ruta""))" } -Cargar 'Leer'
        $r.Error | Should -Not -BeNullOrEmpty
    }
}

Describe 'Libro de análisis en modo Demostracion' -Skip:$SinExcel {
    BeforeAll {
        $script:LibroDemo = Join-Path $CarpetaCorta 'demo.xlsx'
        & $Generador -Modo Demostracion -RutaSalida $LibroDemo 6>$null | Out-Null
        $script:Demo = Get-ConsultasLibro -Ruta $LibroDemo
    }
    It 'tiene las 18 consultas de Power Query' {
        $Demo.Consultas.Count | Should -Be 18
        $Demo.Consultas | Should -Contain 'Registro'
        $Demo.Consultas | Should -Contain 'Consolidado_Operam'
    }
    It 'queda en modo Demostracion' {
        $Demo.Modo | Should -Match '^"Demostracion" meta'
    }
    It 'la tabla <Tabla> tiene <Filas> filas de ejemplo' -ForEach @(
        @{ Tabla = 'tblRegistro'; Filas = 7 }
        @{ Tabla = 'tblSoportes'; Filas = 14 }
        @{ Tabla = 'tblGastosConsolidados'; Filas = 9 }
        @{ Tabla = 'tblErroresLectura'; Filas = 1 }
        @{ Tabla = 'tblConsolidadoOperam'; Filas = 3 }
        @{ Tabla = 'tblAlertas'; Filas = 7 }
        @{ Tabla = 'tblResumenMensual'; Filas = 2 }
        @{ Tabla = 'tblControlOperam'; Filas = 1 }
    ) {
        $Demo.Tablas[$Tabla].Filas | Should -Be $Filas
    }
    It 'el registro tiene los campos acordados con sus nombres visibles' {
        $Demo.Tablas['tblRegistro'].Encabezados | Select-Object -First 9 |
            Should -Be @('ID Legalización', 'Fecha Recepción', 'Colaborador', 'Correo', 'Asunto', 'Nombre Archivo', 'Estado', 'Observaciones', 'Fecha Cierre')
    }
    It 'define la lista de estados para la validación de datos' {
        $Demo.Nombres | Should -Contain 'lstEstados'
    }
}

Describe 'Libro de análisis en modo SharePoint' -Skip:$SinExcel {
    BeforeAll {
        $configPrueba = Join-Path $CarpetaCorta 'Config.Prueba.psd1'
        $texto = [System.IO.File]::ReadAllText((Join-Path $Raiz 'config\Config.Legalizaciones.psd1')).Replace('<PREFIJO_TENANT>', 'contoso')
        [System.IO.File]::WriteAllText($configPrueba, $texto, [System.Text.UTF8Encoding]::new($true))
        $script:LibroSp = Join-Path $CarpetaCorta 'sharepoint.xlsx'
        & $Generador -Modo SharePoint -RutaConfig $configPrueba -RutaSalida $LibroSp 6>$null | Out-Null
        $script:Sp = Get-ConsultasLibro -Ruta $LibroSp
    }
    It 'queda en modo SharePoint, listo para actualizar' {
        $Sp.Modo | Should -Match '^"SharePoint" meta'
    }
    It 'las tablas de consultas quedan vacías pero con sus columnas' {
        foreach ($tabla in 'tblRegistro', 'tblSoportes', 'tblGastosConsolidados', 'tblConsolidadoOperam', 'tblAlertas', 'tblResumenMensual') {
            $Sp.Tablas[$tabla].Filas | Should -Be 0 -Because $tabla
            @($Sp.Tablas[$tabla].Encabezados).Count | Should -BeGreaterThan 1 -Because $tabla
        }
    }
    It 'la tabla manual Control_Operam queda vacía' {
        $Sp.Tablas['tblControlOperam'].Filas | Should -BeLessOrEqual 1
    }
    It 'se niega a generar el modo SharePoint sin tenant configurado' {
        { & $Generador -Modo SharePoint -RutaSalida (Join-Path $CarpetaCorta 'x.xlsx') 6>$null } | Should -Throw '*Tenant.Nombre*'
    }
}
