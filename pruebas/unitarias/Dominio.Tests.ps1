<#
    Pruebas de la capa de DOMINIO (Pester 5).

    Estos ejemplos son el contrato de comportamiento: el flujo de Power Automate debe dar
    exactamente los mismos resultados. El plan de pruebas (pruebas/PLAN-DE-PRUEBAS.md) y el
    catálogo de expresiones (src/3-adaptadores/power-automate/expresiones.md) usan los mismos casos.
#>

BeforeAll {
    $raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
    Import-Module (Join-Path $raiz 'src\1-dominio\Legalizaciones.Dominio.psm1') -Force
    $script:Reglas = Get-ReglasDominio
    function script:Utc([string]$Texto) {
        return [datetime]::Parse($Texto, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
    }
}

Describe 'Reglas del dominio (dominio.json)' {
    It 'tiene los cuatro estados acordados, en orden' {
        @($Reglas.caso.estados) | Should -Be @('Pendiente', 'En revisión', 'Rechazada', 'Aprobada')
    }
    It 'el estado inicial es Pendiente' {
        $Reglas.caso.estadoInicial | Should -Be 'Pendiente'
    }
    It 'cada estado tiene transiciones y carpeta raíz definidas' {
        foreach ($estado in $Reglas.caso.estados) {
            $Reglas.caso.transiciones.PSObject.Properties.Name | Should -Contain $estado
            $Reglas.carpetas.raizPorEstado.PSObject.Properties.Name | Should -Contain $estado
        }
    }
    It 'estados abiertos + finales = todos los estados' {
        (@($Reglas.caso.estadosAbiertos) + @($Reglas.caso.estadosFinales) | Sort-Object) | Should -Be (@($Reglas.caso.estados) | Sort-Object)
    }
    It 'las extensiones aceptadas son .pdf, .xlsx y .xml' {
        Get-ExtensionesPermitidas | Should -Be @('.pdf', '.xlsx', '.xml')
    }
}

Describe 'CU-01 · Reconocer el correo de legalización (condición del desencadenador)' {
    It '"<Asunto>" -> <Esperado>' -ForEach @(
        @{ Asunto = 'Legalización Tarjeta Crédito - Ana Pérez'; Esperado = $true }
        @{ Asunto = 'LEGALIZACIÓN TARJETA CRÉDITO - JUAN DÍAZ'; Esperado = $true }
        @{ Asunto = 'Legalizacion Tarjeta Credito - Luis Gómez'; Esperado = $true }
        @{ Asunto = 'RV: Legalización Tarjeta Crédito - Ana Pérez'; Esperado = $true }
        @{ Asunto = 'Legalización  Tarjeta Crédito - Ana'; Esperado = $true }
        @{ Asunto = 'Legalización de gastos de septiembre'; Esperado = $false }
        @{ Asunto = 'Tarjeta Crédito - legalización'; Esperado = $false }
        @{ Asunto = ''; Esperado = $false }
    ) {
        Test-AsuntoLegalizacion $Asunto | Should -Be $Esperado
    }
}

Describe 'CU-01 · Nombre del colaborador' {
    It 'del asunto "<Asunto>" se toma "<Esperado>"' -ForEach @(
        @{ Asunto = 'Legalización Tarjeta Crédito - Ana Pérez'; Esperado = 'Ana Pérez' }
        @{ Asunto = 'RV: Legalización Tarjeta Crédito - Ana María Pérez-Gómez'; Esperado = 'Ana María Pérez-Gómez' }
        @{ Asunto = 'Legalización Tarjeta Crédito - Ana Pérez - Quincena 1'; Esperado = 'Ana Pérez' }
        @{ Asunto = '  Legalización Tarjeta Crédito -   Ana Pérez  '; Esperado = 'Ana Pérez' }
        @{ Asunto = 'Legalización Tarjeta Crédito'; Esperado = '' }
        @{ Asunto = 'Legalización Tarjeta Crédito - '; Esperado = '' }
        @{ Asunto = 'Legalización Tarjeta Crédito-Ana Pérez'; Esperado = '' }
    ) {
        Get-NombreEnAsunto $Asunto | Should -BeExactly $Esperado
    }

    It 'si el asunto no trae el nombre, usa la parte del correo antes de la @ en minúsculas' {
        Get-NombreColaborador -Asunto 'Legalización Tarjeta Crédito' -Remitente ' Ana.Perez@Emobility.com ' | Should -BeExactly 'ana.perez'
    }
    It 'si el asunto trae el nombre, lo usa aunque haya remitente' {
        Get-NombreColaborador -Asunto 'Legalización Tarjeta Crédito - Ana Pérez' -Remitente 'otra@emobility.com' | Should -BeExactly 'Ana Pérez'
    }
}

Describe 'CU-02 · Clasificar adjuntos por extensión' {
    It '"<Nombre>" -> tipo <Esperado>' -ForEach @(
        @{ Nombre = 'Factura_123.PDF'; Esperado = 'factura' }
        @{ Nombre = 'Gastos quincena.xlsx'; Esperado = 'comprobante' }
        @{ Nombre = 'cfdi.XML'; Esperado = 'xml' }
        @{ Nombre = 'Gastos.xls'; Esperado = $null }
        @{ Nombre = 'Gastos.xlsm'; Esperado = $null }
        @{ Nombre = 'foto_factura.jpg'; Esperado = $null }
        @{ Nombre = 'archivo'; Esperado = $null }
        @{ Nombre = 'factura.pdf.exe'; Esperado = $null }
    ) {
        Get-TipoDocumento $Nombre | Should -Be $Esperado
    }
    It 'la extensión de "archivo" (sin punto) es ".archivo", igual que en el flujo' {
        Get-ExtensionArchivo 'archivo' | Should -BeExactly '.archivo'
    }
}

Describe 'CU-02 · Validar la documentación' {
    It 'completa con factura y comprobante' {
        $r = Test-Documentacion -Adjuntos @('factura.pdf', 'gastos.xlsx') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeTrue
        $r.Hallazgos | Should -BeNullOrEmpty
        $r.Soportados | Should -Be @('factura.pdf', 'gastos.xlsx')
    }
    It 'el XML es opcional' {
        $r = Test-Documentacion -Adjuntos @('cfdi.xml', 'gastos.xlsx', 'factura.pdf') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeTrue
        $r.Soportados | Should -Be @('factura.pdf', 'gastos.xlsx', 'cfdi.xml')
    }
    It 'sin factura: incompleta y lo dice' {
        $r = Test-Documentacion -Adjuntos @('gastos.xlsx') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeFalse
        $r.Hallazgos | Should -Be @($Reglas.mensajes.faltaFactura)
    }
    It 'sin comprobante: incompleta y lo dice' {
        $r = Test-Documentacion -Adjuntos @('factura.pdf') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Hallazgos | Should -Be @($Reglas.mensajes.faltaComprobante)
    }
    It 'sin adjuntos: faltan ambos, en ese orden' {
        $r = Test-Documentacion -Adjuntos @() -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeFalse
        $r.Hallazgos | Should -Be @($Reglas.mensajes.faltaFactura, $Reglas.mensajes.faltaComprobante)
    }
    It 'factura en JPG: falta la factura y el JPG se informa como no soportado' {
        $r = Test-Documentacion -Adjuntos @('factura.jpg', 'gastos.xlsx') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeFalse
        $r.NoSoportados | Should -Be @('factura.jpg')
        $r.Hallazgos[0] | Should -BeExactly $Reglas.mensajes.faltaFactura
        $r.Hallazgos[1] | Should -BeExactly 'Estos archivos no se guardaron porque su formato no es válido: factura.jpg. Formatos aceptados: .pdf, .xlsx, .xml.'
    }
    It 'un archivo extra no soportado no impide que esté completa' {
        $r = Test-Documentacion -Adjuntos @('factura.pdf', 'gastos.xlsx', 'nota.docx') -Asunto 'Legalización Tarjeta Crédito - Ana'
        $r.Completa | Should -BeTrue
        $r.NoSoportados | Should -Be @('nota.docx')
        $r.Hallazgos.Count | Should -Be 1
    }
    It 'avisa si el asunto no trae el nombre' {
        $r = Test-Documentacion -Adjuntos @('factura.pdf', 'gastos.xlsx') -Asunto 'Legalización Tarjeta Crédito'
        $r.Completa | Should -BeTrue
        $r.Hallazgos | Should -Be @($Reglas.mensajes.asuntoSinNombre)
    }
}

Describe 'CU-01 · Identificador, fechas y carpetas' {
    It 'Id con mes local y consecutivo de 5 dígitos: <Esperado>' -ForEach @(
        @{ Fecha = '2026-09-15 10:00'; Consecutivo = 123; Esperado = 'LEG-202609-00123' }
        @{ Fecha = '2026-01-02 08:00'; Consecutivo = 7; Esperado = 'LEG-202601-00007' }
        @{ Fecha = '2026-09-15 10:00'; Consecutivo = 123456; Esperado = 'LEG-202609-123456' }
    ) {
        New-IdLegalizacion -FechaRecepcionLocal ([datetime]$Fecha) -Consecutivo $Consecutivo | Should -BeExactly $Esperado
    }

    It 'lee año y mes desde el Id' {
        $periodo = Get-PeriodoDesdeId 'LEG-202608-00031'
        $periodo.Anio | Should -BeExactly '2026'
        $periodo.Mes | Should -BeExactly '08'
    }

    It 'un correo recibido el 1/sep a las 03:30 UTC es del 31/ago en México (carpeta 08)' {
        $local = ConvertTo-FechaLocal -FechaUtc (Utc '2026-09-01T03:30:00Z') -ZonaHoraria 'Central Standard Time (Mexico)'
        $local.ToString('yyyy-MM-dd HH:mm') | Should -BeExactly '2026-08-31 21:30'
    }
    It 'y en Colombia a las 04:30 UTC también es 31/ago' {
        $local = ConvertTo-FechaLocal -FechaUtc (Utc '2026-09-01T04:30:00Z') -ZonaHoraria 'SA Pacific Standard Time'
        $local.ToString('yyyy-MM-dd HH:mm') | Should -BeExactly '2026-08-31 23:30'
    }

    It 'nombre de carpeta seguro: "<Texto>" -> "<Esperado>"' -ForEach @(
        @{ Texto = 'Ana Pérez'; Esperado = 'Ana Pérez' }
        @{ Texto = 'Ana P.'; Esperado = 'Ana P' }
        @{ Texto = 'Juan "El Rápido" Díaz'; Esperado = 'Juan El Rápido Díaz' }
        @{ Texto = 'María/José: #1'; Esperado = 'MaríaJosé 1' }
        @{ Texto = '  ???  '; Esperado = 'Sin_nombre' }
        @{ Texto = ''; Esperado = 'Sin_nombre' }
    ) {
        ConvertTo-NombreCarpeta $Texto | Should -BeExactly $Esperado
    }
    It 'recorta nombres largos a 60 caracteres sin espacio final' {
        $largo = ('Nombre ' * 12)
        $carpeta = ConvertTo-NombreCarpeta $largo
        $carpeta.Length | Should -BeLessOrEqual 60
        $carpeta | Should -Not -Match '\s$'
    }

    It 'carpeta del caso para <Estado>: <Esperado>' -ForEach @(
        @{ Estado = 'Pendiente'; Esperado = 'Pendientes/2026/08/Ana Pérez' }
        @{ Estado = 'En revisión'; Esperado = 'Pendientes/2026/08/Ana Pérez' }
        @{ Estado = 'Aprobada'; Esperado = 'Aprobadas/2026/08/Ana Pérez' }
        @{ Estado = 'Rechazada'; Esperado = 'Rechazadas/2026/08/Ana Pérez' }
    ) {
        Get-RutaCarpetaCaso -Estado $Estado -FechaRecepcionLocal ([datetime]'2026-08-31 21:30') -Colaborador 'Ana Pérez' | Should -BeExactly $Esperado
    }
}

Describe 'CU-05 · Transiciones de estado' {
    It '<Desde> -> <Hacia>: <Esperado>' -ForEach @(
        @{ Desde = 'Pendiente'; Hacia = 'En revisión'; Esperado = $true }
        @{ Desde = 'Pendiente'; Hacia = 'Aprobada'; Esperado = $true }
        @{ Desde = 'Pendiente'; Hacia = 'Rechazada'; Esperado = $true }
        @{ Desde = 'En revisión'; Hacia = 'Pendiente'; Esperado = $true }
        @{ Desde = 'En revisión'; Hacia = 'Aprobada'; Esperado = $true }
        @{ Desde = 'Aprobada'; Hacia = 'Pendiente'; Esperado = $false }
        @{ Desde = 'Rechazada'; Hacia = 'Aprobada'; Esperado = $false }
        @{ Desde = 'Aprobada'; Hacia = 'Aprobada'; Esperado = $true }
    ) {
        Test-TransicionEstado -Desde $Desde -Hacia $Hacia | Should -Be $Esperado
    }
    It 'rechaza estados que no existen' {
        { Test-TransicionEstado -Desde 'Pendiente' -Hacia 'Cerrada' } | Should -Throw
    }
}

Describe 'CU-03 · Posibles duplicados (heurística)' {
    BeforeAll {
        $script:Previos = @(
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00001'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-10T15:00:00Z'); Estado = 'Pendiente'; NombreArchivo = "factura_001.pdf`ngastos.xlsx" }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00002'; Correo = 'ANA@EMOBILITY.COM'; FechaRecepcionUtc = (Utc '2026-09-12T15:00:00Z'); Estado = 'Aprobada'; NombreArchivo = 'FACTURA_001.PDF' }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202607-00003'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-07-01T15:00:00Z'); Estado = 'Pendiente'; NombreArchivo = 'factura_001.pdf' }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00004'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-15T15:00:00Z'); Estado = 'Rechazada'; NombreArchivo = 'factura_001.pdf' }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00005'; Correo = 'luis@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-15T15:00:00Z'); Estado = 'Pendiente'; NombreArchivo = 'factura_001.pdf' }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00006'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-18T15:00:00Z'); Estado = 'Pendiente'; NombreArchivo = "otra.pdf`ngastos.xlsx" }
            [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00007'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-18T16:00:00Z'); Estado = 'En revisión'; NombreArchivo = "factura_001.pdf`r`ngastos.xlsx" }
        )
    }
    It 'encuentra solo los casos del mismo remitente, dentro de 45 días, no rechazados y con el mismo PDF' {
        $ids = Find-PosiblesDuplicados -NuevosArchivos @('Factura_001.pdf', 'gastos.xlsx') -Remitente 'ana@emobility.com' -FechaRecepcionUtc (Utc '2026-09-20T15:00:00Z') -CasosPrevios $Previos
        $ids | Should -Be @('LEG-202609-00001', 'LEG-202609-00002', 'LEG-202609-00007')
    }
    It 'el Excel quincenal no cuenta: suele llamarse igual en cada envío' {
        $ids = Find-PosiblesDuplicados -NuevosArchivos @('gastos.xlsx') -Remitente 'ana@emobility.com' -FechaRecepcionUtc (Utc '2026-09-20T15:00:00Z') -CasosPrevios $Previos
        $ids | Should -BeNullOrEmpty
    }
    It 'el XML sí cuenta' {
        $previo = [pscustomobject]@{ IdLegalizacion = 'LEG-202609-00010'; Correo = 'ana@emobility.com'; FechaRecepcionUtc = (Utc '2026-09-19T15:00:00Z'); Estado = 'Pendiente'; NombreArchivo = 'CFDI_A1.xml' }
        Find-PosiblesDuplicados -NuevosArchivos @('cfdi_a1.XML') -Remitente 'ana@emobility.com' -FechaRecepcionUtc (Utc '2026-09-20T15:00:00Z') -CasosPrevios @($previo) | Should -Be @('LEG-202609-00010')
    }
    It 'sin casos previos no hay duplicados' {
        Find-PosiblesDuplicados -NuevosArchivos @('factura.pdf') -Remitente 'ana@emobility.com' -FechaRecepcionUtc (Utc '2026-09-20T15:00:00Z') -CasosPrevios @() | Should -BeNullOrEmpty
    }
    It 'arma el aviso para Contabilidad' {
        Get-MensajePosibleDuplicado -Casos @('LEG-202609-00001', 'LEG-202609-00002') |
            Should -BeExactly '[Interno] Posible duplicado de LEG-202609-00001, LEG-202609-00002: mismo remitente y mismo nombre de archivo en los últimos 45 días.'
    }
    It 'el aviso de duplicado es una nota interna' {
        (Get-MensajePosibleDuplicado -Casos @('LEG-1')).StartsWith($Reglas.observaciones.prefijoNotaInterna) | Should -BeTrue
    }
}

Describe 'CU-05 · Observaciones que ve el colaborador al cerrar' {
    It 'quita las notas internas y las líneas vacías' {
        $observaciones = "Falta el comprobante quincenal en Excel (.xlsx).`r`n[Interno] Posible duplicado de LEG-202609-00001: mismo remitente...`n`nFactura ilegible, envíe una copia legible."
        Get-ObservacionesParaColaborador $observaciones | Should -Be @('Falta el comprobante quincenal en Excel (.xlsx).', 'Factura ilegible, envíe una copia legible.')
    }
    It 'sin observaciones no devuelve nada' {
        Get-ObservacionesParaColaborador '' | Should -BeNullOrEmpty
    }
    It 'una observación solo interna no se muestra' {
        Get-ObservacionesParaColaborador '[Interno] ERROR TÉCNICO al procesar el correo' | Should -BeNullOrEmpty
    }
}
